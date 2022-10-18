use crossbeam_queue::SegQueue;
use pyo3::exceptions::{PyOSError, PyRuntimeError, PyTypeError};
use pyo3::intern;
use pyo3::prelude::*;
use std::error::Error;
use std::fmt;
use std::sync::Arc;
use windows::{
    core::HSTRING, Foundation::Metadata::ApiInformation, Foundation::TypedEventHandler,
    Media::Core::MediaSource, Media::Playback::*, Media::SpeechSynthesis::*, Storage::StorageFile,
    Storage::Streams::InMemoryRandomAccessStream,
};

pub type NeosynthResult<T> = Result<T, NeosynthError>;
pub use NeosynthError::{OperationError, RuntimeError};

#[derive(Debug)]
pub enum NeosynthError {
    RuntimeError(String, i32),
    OperationError(String),
}

impl Error for NeosynthError {}

impl fmt::Display for NeosynthError {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        let err_message = match self {
            RuntimeError(msg, code) => format!("Windows error: {} Code: {}.", msg, code),
            OperationError(msg) => format!("Error: {}", msg),
        };
        write!(f, "{}", err_message)
    }
}

impl From<windows::core::Error> for NeosynthError {
    fn from(error: windows::core::Error) -> Self {
        RuntimeError(error.message().to_string(), error.code().0)
    }
}

impl From<NeosynthError> for PyErr {
    fn from(error: NeosynthError) -> Self {
        match error {
            RuntimeError(msg, code) => PyOSError::new_err((msg, code)),
            OperationError(msg) => PyRuntimeError::new_err(msg),
        }
    }
}

#[derive(Debug, Clone)]
#[pyclass]
pub enum SynthState {
    Ready = 0,
    Busy = 1,
    Paused = 2,
}

impl Default for SynthState {
    fn default() -> Self {
        SynthState::Ready
    }
}

impl From<MediaPlaybackState> for SynthState {
    fn from(player_state: MediaPlaybackState) -> Self {
        match player_state {
            MediaPlaybackState::Buffering
            | MediaPlaybackState::Opening
            | MediaPlaybackState::Playing => SynthState::Busy,
            MediaPlaybackState::Paused => SynthState::Paused,
            _ => SynthState::Ready,
        }
    }
}

#[derive(Clone)]
pub enum SpeechElement {
    Text(String),
    Ssml(String),
    Bookmark(String),
    Audio(String),
}

#[derive(Clone)]
#[pyclass(subclass)]
pub struct SpeechUtterance {
    content: Vec<SpeechElement>,
}

impl Default for SpeechUtterance {
    fn default() -> Self {
        Self::new()
    }
}

#[pymethods]
impl SpeechUtterance {
    #[new]
    pub fn new() -> Self {
        Self {
            content: Vec::new(),
        }
    }
    #[pyo3(text_signature = "($self, text: str)")]
    fn add_text(&mut self, text: String) {
        self.content.push(SpeechElement::Text(text));
    }
    #[pyo3(text_signature = "($self, ssml: str)")]
    fn add_ssml(&mut self, ssml: String) {
        self.content.push(SpeechElement::Ssml(ssml));
    }
    #[pyo3(text_signature = "($self, bookmark: str)")]
    fn add_bookmark(&mut self, bookmark: String) {
        self.content.push(SpeechElement::Bookmark(bookmark));
    }
    #[pyo3(text_signature = "($self, audio_path: str)")]
    fn add_audio(&mut self, audio_path: String) {
        self.content.push(SpeechElement::Audio(audio_path));
    }
    #[pyo3(text_signature = "($self, utterance: neosynth.SpeechUtterance)")]
    fn add_utterance(&mut self, utterance: &mut Self) {
        self.content.append(&mut utterance.content);
    }
}

#[derive(Debug)]
#[pyclass(frozen)]
pub struct VoiceInfo {
    #[pyo3(get)]
    pub id: String,
    #[pyo3(get)]
    pub language: String,
    #[pyo3(get)]
    pub name: String,
    voice: VoiceInformation,
}

impl From<VoiceInformation> for VoiceInfo {
    fn from(vinfo: VoiceInformation) -> Self {
        VoiceInfo {
            id: vinfo.Id().unwrap().to_string(),
            language: vinfo.Language().unwrap().to_string(),
            name: vinfo.DisplayName().unwrap().to_string(),
            voice: vinfo,
        }
    }
}

impl From<&VoiceInfo> for VoiceInformation {
    fn from(vinfo: &VoiceInfo) -> Self {
        vinfo.voice.clone()
    }
}

pub trait NsEventSink {
    fn on_state_changed(&self, new_state: SynthState);
    fn on_bookmark_reached(&self, bookmark: String);
}

pub struct PyEventSinkWrapper {
    py_event_sink: PyObject,
}

impl PyEventSinkWrapper {
    fn new(py_event_sink: PyObject) -> Self {
        Self { py_event_sink }
    }
}

impl NsEventSink for PyEventSinkWrapper {
    fn on_state_changed(&self, new_state: SynthState) {
        Python::with_gil(|py| {
            self.py_event_sink
                .call_method1(py, "on_state_changed", (new_state,))
                .ok();
        });
    }
    fn on_bookmark_reached(&self, bookmark: String) {
        Python::with_gil(|py| {
            self.py_event_sink
                .call_method1(py, "on_bookmark_reached", (bookmark,))
                .ok();
        });
    }
}

struct SpeechMixer<T>
where
    T: NsEventSink + std::marker::Send + 'static,
{
    synthesizer: SpeechSynthesizer,
    player: MediaPlayer,
    speech_queue: SegQueue<SpeechElement>,
    event_sink: T,
}

impl<T> SpeechMixer<T>
where
    T: NsEventSink + std::marker::Send + 'static,
{
    pub fn new(event_sink: T) -> NeosynthResult<Self> {
        Ok(Self {
            synthesizer: SpeechSynthesizer::new()?,
            player: MediaPlayer::new()?,
            speech_queue: SegQueue::new(),
            event_sink,
        })
    }

    pub fn get_state(&self) -> NeosynthResult<SynthState> {
        Ok(self.player.PlaybackSession()?.PlaybackState()?.into())
    }

    pub fn speak_content(&self, text: &str, is_ssml: bool) -> NeosynthResult<()> {
        let stream = self.generate_speech_stream(text, is_ssml)?;
        self.player.SetSource(&MediaSource::CreateFromStream(
            &stream,
            &stream.ContentType()?,
        )?)?;
        Ok(())
    }

    fn generate_speech_stream(
        &self,
        text: &str,
        is_ssml: bool,
    ) -> NeosynthResult<SpeechSynthesisStream> {
        let output = if is_ssml {
            self.synthesizer
                .SynthesizeSsmlToStreamAsync(&HSTRING::from(text))?
                .get()?
        } else {
            self.synthesizer
                .SynthesizeTextToStreamAsync(&HSTRING::from(text))?
                .get()?
        };
        Ok(output)
    }

    pub fn process_speech_element(&self, element: SpeechElement) -> NeosynthResult<()> {
        match element {
            SpeechElement::Text(text) => self.speak_content(&text, false),
            SpeechElement::Ssml(ssml) => self.speak_content(&ssml, true),
            SpeechElement::Bookmark(bookmark) => {
                self.event_sink.on_bookmark_reached(bookmark);
                self.process_queue()
            }
            SpeechElement::Audio(filename) => {
                let audiofile =
                    StorageFile::GetFileFromPathAsync(&HSTRING::from(filename))?.get()?;
                self.player
                    .SetSource(&MediaSource::CreateFromStorageFile(&audiofile)?)?;
                Ok(())
            }
        }
    }

    fn process_queue(&self) -> NeosynthResult<()> {
        match self.speech_queue.pop() {
            Some(elem) => self.process_speech_element(elem),
            None => {
                self.event_sink.on_state_changed(SynthState::Ready);
                Ok(())
            }
        }
    }

    pub fn speak<I>(&self, utterance: I) -> NeosynthResult<()>
    where
        I: IntoIterator<Item = SpeechElement>,
    {
        utterance
            .into_iter()
            .for_each(|elem| self.speech_queue.push(elem));
        match self.get_state()? {
            SynthState::Ready => match self.speech_queue.pop() {
                Some(element) => self.process_speech_element(element),
                None => Ok(()),
            },
            _ => Ok(()),
        }
    }
}

#[pyclass(subclass, frozen)]
pub struct Neosynth {
    mixer: Arc<SpeechMixer<PyEventSinkWrapper>>,
}

impl Neosynth {
    pub fn new(event_sink_wrapper: PyEventSinkWrapper) -> NeosynthResult<Self> {
        let instance = Self {
            mixer: Arc::new(SpeechMixer::new(event_sink_wrapper)?),
        };
        instance.initialize()?;
        Ok(instance)
    }
    fn initialize(&self) -> NeosynthResult<()> {
        self.mixer.player.SetAutoPlay(true)?;
        // Remove extended silence at the end of each speech utterance
        if ApiInformation::IsApiContractPresentByMajorAndMinor(
            &HSTRING::from("Windows.Foundation.UniversalApiContract"),
            6,
            0,
        )? {
            self.mixer
                .synthesizer
                .Options()?
                .SetAppendedSilence(SpeechAppendedSilence::Min)?;
        };
        self.register_player_events()
    }

    fn register_player_events(&self) -> NeosynthResult<()> {
        let mixer = Arc::clone(&self.mixer);
        self.mixer
            .player
            .MediaEnded(&TypedEventHandler::<MediaPlayer, _>::new(move |_, _| {
                mixer.process_queue().ok();
                Ok(())
            }))?;
        let mixer = Arc::clone(&self.mixer);
        self.mixer.player.MediaFailed(&TypedEventHandler::<
            MediaPlayer,
            MediaPlayerFailedEventArgs,
        >::new(move |_, _| {
            mixer.process_queue().ok();
            Ok(())
        }))?;
        let mixer = Arc::clone(&self.mixer);
        self.mixer
            .player
            .PlaybackSession()?
            .PlaybackStateChanged(&TypedEventHandler::<MediaPlaybackSession, _>::new(
                move |_, _| {
                    mixer
                        .event_sink
                        .on_state_changed(mixer.get_state().unwrap());
                    Ok(())
                },
            ))?;
        Ok(())
    }

    fn is_prosody_supported() -> NeosynthResult<bool> {
        Ok(ApiInformation::IsApiContractPresentByMajorAndMinor(
            &HSTRING::from("Windows.Foundation.UniversalApiContract"),
            5,
            0,
        )?)
    }
}

#[pymethods]
impl Neosynth {
    #[new]
    pub fn py_init(py: Python<'_>, event_sink: PyObject) -> PyResult<Self> {
        let obj: &PyAny = event_sink.as_ref(py);
        if (!obj.hasattr(intern!(py, "on_state_changed"))?)
            || (!obj.hasattr(intern!(py, "on_bookmark_reached"))?)
        {
            Err(PyTypeError::new_err(
                "The provided object does not have the required method handlers.",
            ))
        } else {
            Ok(Self::new(PyEventSinkWrapper::new(event_sink))?)
        }
    }

    /// Get the current state of the synthesizer
    #[pyo3(text_signature = "($self) -> neosynth.SynthState")]
    pub fn get_state(&self) -> NeosynthResult<SynthState> {
        self.mixer.get_state()
    }

    /// Get the current volume
    #[pyo3(text_signature = "($self) -> float")]
    pub fn get_volume(&self) -> NeosynthResult<f64> {
        Ok(self.mixer.player.Volume()? * 100f64)
    }

    /// Set the current volume
    #[pyo3(text_signature = "($self, volume: float)")]
    pub fn set_volume(&self, volume: f64) -> NeosynthResult<()> {
        Ok(self.mixer.player.SetVolume(volume / 100f64)?)
    }
    /// Get the current speaking rate
    #[pyo3(text_signature = "($self) -> float")]
    pub fn get_rate(&self) -> NeosynthResult<f64> {
        if !Self::is_prosody_supported()? {
            Ok(-1.0)
        } else {
            Ok(self.mixer.synthesizer.Options()?.SpeakingRate()? / 0.06)
        }
    }
    /// Set the current speaking rate
    #[pyo3(text_signature = "($self, rate: float)")]
    pub fn set_rate(&self, value: f64) -> NeosynthResult<()> {
        if Self::is_prosody_supported()? {
            Ok(self
                .mixer
                .synthesizer
                .Options()?
                .SetSpeakingRate(value * 0.06)?)
        } else {
            Err(NeosynthError::OperationError(
                "The current version of OneCore synthesizer does not support the prosody option"
                    .to_string(),
            ))
        }
    }
    /// Get the current voice
    #[pyo3(text_signature = "($self) -> neosynth.VoiceInfo")]
    pub fn get_voice(&self) -> NeosynthResult<VoiceInfo> {
        Ok(self.mixer.synthesizer.Voice()?.into())
    }
    /// Set the current voice
    #[pyo3(text_signature = "($self, voice: neosynth.VoiceInfo)")]
    pub fn set_voice(&self, voice: &VoiceInfo) -> NeosynthResult<()> {
        Ok(self
            .mixer
            .synthesizer
            .SetVoice(&VoiceInformation::from(voice))?)
    }
    /// Get the current voice's string representation
    #[pyo3(text_signature = "($self) -> str")]
    pub fn get_voice_str(&self) -> NeosynthResult<String> {
        Ok(self.get_voice()?.id)
    }
    /// Set the current voice using a previously obtained string representation of a voice
    #[pyo3(text_signature = "($self, voice_str: str)")]
    pub fn set_voice_str(&self, id: String) -> NeosynthResult<()> {
        let voice = Self::get_voices()?.into_iter().find(|v| v.id == id);
        match voice {
            Some(v) => self.set_voice(&v),
            None => Err(OperationError("Invalid voice token given".to_string())),
        }
    }
    /// Get a list of installed voices
    #[staticmethod]
    #[pyo3(text_signature = "() -> list[neosynth.VoiceInfo]")]
    pub fn get_voices() -> NeosynthResult<Vec<VoiceInfo>> {
        let voices: Vec<VoiceInfo> = SpeechSynthesizer::AllVoices()?
            .into_iter()
            .map(VoiceInfo::from)
            .collect();
        Ok(voices)
    }
    /// Speak a neosynth.SpeechUtterance
    #[pyo3(text_signature = "($self, utterance: neosynth.SpeechUtterance)")]
    pub fn speak(&self, utterance: SpeechUtterance) -> NeosynthResult<()> {
        self.mixer.speak(utterance.content)
    }
    /// Pause the speech
    #[pyo3(text_signature = "($self)")]
    pub fn pause(&self) -> NeosynthResult<()> {
        Ok(self.mixer.player.Pause()?)
    }
    /// Resume the speech
    #[pyo3(text_signature = "($self)")]
    pub fn resume(&self) -> NeosynthResult<()> {
        Ok(self.mixer.player.Play()?)
    }
    /// Stop the speech
    #[pyo3(text_signature = "($self)")]
    pub fn stop(&self) -> NeosynthResult<()> {
        self.pause()?;
        // Drop any queued elements
        loop {
            if self.mixer.speech_queue.pop().is_none() {
                break;
            }
        }
        let empty_stream = InMemoryRandomAccessStream::new()?;
        self.mixer.player.SetSource(&MediaSource::CreateFromStream(
            &empty_stream,
            &HSTRING::from(""),
        )?)?;
        self.mixer.process_queue()
    }
}

/// A wrapper around Windows OneCoreSynthesizer
#[pymodule]
fn neosynth(_py: Python<'_>, m: &PyModule) -> PyResult<()> {
    m.add_class::<Neosynth>()?;
    m.add_class::<SynthState>()?;
    m.add_class::<SpeechUtterance>()?;
    m.add_class::<VoiceInfo>()?;
    Ok(())
}
