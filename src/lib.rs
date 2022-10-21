use crossbeam_queue::SegQueue;
use pyo3::exceptions::{PyOSError, PyRuntimeError, PyTypeError};
use pyo3::intern;
use pyo3::prelude::*;
use std::error::Error;
use std::fmt;
use std::sync::{Arc, RwLock};
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

#[pyclass]
#[derive(Default, Eq, PartialEq, Copy, Clone)]
pub enum SynthState {
    #[default]
    Ready = 0,
    Busy = 1,
    Paused = 2,
}

#[pymethods]
impl SynthState {
    fn __hash__(&self) -> PyResult<isize> {
        let evalue = match self {
            SynthState::Ready => 0,
            SynthState::Busy => 1,
            SynthState::Paused => 2,
        };
        Ok(evalue)
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

#[pyclass(subclass)]
#[derive(Default, Clone)]
pub struct SpeechUtterance(Vec<SpeechElement>);

#[pymethods]
impl SpeechUtterance {
    #[new]
    pub fn new() -> Self {
        Default::default()
    }
    #[pyo3(text_signature = "($self, text: str)")]
    fn add_text(&mut self, text: String) {
        self.0.push(SpeechElement::Text(text));
    }
    #[pyo3(text_signature = "($self, ssml: str)")]
    fn add_ssml(&mut self, ssml: String) {
        self.0.push(SpeechElement::Ssml(ssml));
    }
    #[pyo3(text_signature = "($self, bookmark: str)")]
    fn add_bookmark(&mut self, bookmark: String) {
        self.0.push(SpeechElement::Bookmark(bookmark));
    }
    #[pyo3(text_signature = "($self, audio_path: str)")]
    fn add_audio(&mut self, audio_path: String) {
        self.0.push(SpeechElement::Audio(audio_path));
    }
    #[pyo3(text_signature = "($self, utterance: neosynth.SpeechUtterance)")]
    fn add_utterance(&mut self, utterance: &mut Self) {
        self.0.append(&mut utterance.0);
    }
}

#[pyclass(frozen)]
#[derive(Debug)]
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

pub struct PyEventSinkWrapper(PyObject);

impl PyEventSinkWrapper {
    fn new(py_event_sink: PyObject) -> Self {
        Self(py_event_sink)
    }
}

impl NsEventSink for PyEventSinkWrapper {
    fn on_state_changed(&self, new_state: SynthState) {
        Python::with_gil(|py| {
            self.0
                .call_method1(py, "on_state_changed", (new_state,))
                .ok();
        });
    }
    fn on_bookmark_reached(&self, bookmark: String) {
        Python::with_gil(|py| {
            self.0
                .call_method1(py, "on_bookmark_reached", (bookmark,))
                .ok();
        });
    }
}

pub struct NeoMediaPlayer(MediaPlayer);

impl NeoMediaPlayer {
    fn new() -> NeosynthResult<Self> {
        let win_player = MediaPlayer::new()?;
        win_player.SetAudioCategory(MediaPlayerAudioCategory::Speech)?;
        Ok(Self(win_player))
    }
    pub fn get_playback_state(&self) -> NeosynthResult<MediaPlaybackState> {
        Ok(self.0.PlaybackSession()?.PlaybackState()?)
    }
    pub fn get_volume(&self) -> NeosynthResult<f64> {
        Ok(self.0.Volume()? * 100f64)
    }
    pub fn set_volume(&self, volume: f64) -> NeosynthResult<()> {
        Ok(self.0.SetVolume(volume / 100f64)?)
    }
    fn set_speech_stream_source(&self, stream: SpeechSynthesisStream) -> NeosynthResult<()> {
        self.0.SetSource(&MediaSource::CreateFromStream(
            &stream,
            &stream.ContentType()?,
        )?)?;
        Ok(())
    }
    fn set_file_source(&self, file_path: String) -> NeosynthResult<()> {
        let audiofile = StorageFile::GetFileFromPathAsync(&HSTRING::from(file_path))?.get()?;
        self.0
            .SetSource(&MediaSource::CreateFromStorageFile(&audiofile)?)?;
        Ok(())
    }
    fn play(&self) -> NeosynthResult<()> {
        self.0.Play()?;
        Ok(())
    }
    fn pause(&self) -> NeosynthResult<()> {
        self.0.Pause()?;
        Ok(())
    }
    fn resume(&self) -> NeosynthResult<()> {
        if self.get_playback_state()? == MediaPlaybackState::Paused {
            self.0.Play()?;
        }
        Ok(())
    }
    fn stop(&self) -> NeosynthResult<()> {
        self.pause()?;
        let empty_stream = InMemoryRandomAccessStream::new()?;
        self.0.SetSource(&MediaSource::CreateFromStream(
            &empty_stream,
            &HSTRING::from(""),
        )?)?;
        Ok(())
    }
}

struct SpeechMixer<T>
where
    T: NsEventSink + std::marker::Send + std::marker::Sync + 'static,
{
    synthesizer: SpeechSynthesizer,
    player: NeoMediaPlayer,
    state: RwLock<SynthState>,
    speech_queue: SegQueue<SpeechElement>,
    event_sink: T,
}

impl<T> SpeechMixer<T>
where
    T: NsEventSink + std::marker::Send + std::marker::Sync + 'static,
{
    pub fn new(event_sink: T) -> NeosynthResult<Self> {
        Ok(Self {
            synthesizer: SpeechSynthesizer::new()?,
            player: NeoMediaPlayer::new()?,
            state: RwLock::new(Default::default()),
            speech_queue: SegQueue::new(),
            event_sink,
        })
    }

    pub fn get_state(&self) -> NeosynthResult<SynthState> {
        Ok(*self.state.read().unwrap())
    }

    pub fn set_state(&self, state: SynthState) -> NeosynthResult<()> {
        if self.get_state()? != state {
            let mut wst = self.state.write().unwrap();
            *wst = state;
            self.event_sink.on_state_changed(state);
        }
        Ok(())
    }

    pub fn speak_content(&self, text: &str, is_ssml: bool) -> NeosynthResult<()> {
        let stream = self.generate_speech_stream(text, is_ssml)?;
        self.player.set_speech_stream_source(stream)?;
        self.player.play()?;
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
            SpeechElement::Text(text) => self.speak_content(&text, false)?,
            SpeechElement::Ssml(ssml) => self.speak_content(&ssml, true)?,
            SpeechElement::Bookmark(bookmark) => {
                self.event_sink.on_bookmark_reached(bookmark);
                self.process_queue()?;
            }
            SpeechElement::Audio(filename) => self.player.set_file_source(filename)?,
        };
        Ok(())
    }

    fn process_queue(&self) -> NeosynthResult<()> {
        match self.speech_queue.pop() {
            Some(elem) => self.process_speech_element(elem),
            None => {
                self.set_state(SynthState::Ready)?;
                Ok(())
            }
        }
    }

    pub fn speak<I>(&self, utterance: I) -> NeosynthResult<()>
    where
        I: IntoIterator<Item = SpeechElement>,
    {
        self.clear_speech_queue()?;
        utterance
            .into_iter()
            .for_each(|elem| self.speech_queue.push(elem));
        self.process_queue()
    }
    pub fn clear_speech_queue(&self) -> NeosynthResult<()> {
        loop {
            if self.speech_queue.pop().is_none() {
                break;
            }
        }
        Ok(())
    }
}

#[pyclass(subclass, frozen)]
pub struct Neosynth(Arc<SpeechMixer<PyEventSinkWrapper>>);

impl Neosynth {
    pub fn new(event_sink_wrapper: PyEventSinkWrapper) -> NeosynthResult<Self> {
        let instance = Self(Arc::new(SpeechMixer::new(event_sink_wrapper)?));
        instance.initialize()?;
        Ok(instance)
    }

    fn initialize(&self) -> NeosynthResult<()> {
        // Remove extended silence at the end of each speech utterance
        if ApiInformation::IsApiContractPresentByMajorAndMinor(
            &HSTRING::from("Windows.Foundation.UniversalApiContract"),
            6,
            0,
        )? {
            self.0
                .synthesizer
                .Options()?
                .SetAppendedSilence(SpeechAppendedSilence::Min)?;
            self.0
                .synthesizer
                .Options()?
                .SetPunctuationSilence(SpeechPunctuationSilence::Min)?;
        }
        self.register_events()?;
        Ok(())
    }

    fn register_events(&self) -> NeosynthResult<()> {
        let mixer = Arc::clone(&self.0);
        self.0
            .player
            .0
            .MediaEnded(&TypedEventHandler::<MediaPlayer, _>::new(move |_, _| {
                mixer.process_queue().ok();
                Ok(())
            }))?;
        let mixer = Arc::clone(&self.0);
        self.0
            .player
            .0
            .MediaFailed(&TypedEventHandler::<MediaPlayer, _>::new(move |_, _| {
                mixer.process_queue().ok();
                Ok(())
            }))?;
        Ok(())
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
    /// Indicates if the prosody option is supported
    #[staticmethod]
    pub fn is_prosody_supported() -> NeosynthResult<bool> {
        Ok(ApiInformation::IsApiContractPresentByMajorAndMinor(
            &HSTRING::from("Windows.Foundation.UniversalApiContract"),
            5,
            0,
        )?)
    }
    /// Get the current state of the synthesizer
    #[pyo3(text_signature = "($self) -> neosynth.SynthState")]
    pub fn get_state(&self) -> NeosynthResult<SynthState> {
        self.0.get_state()
    }

    /// Get the current volume
    #[pyo3(text_signature = "($self) -> float")]
    pub fn get_volume(&self) -> NeosynthResult<f64> {
        self.0.player.get_volume()
    }

    /// Set the current volume
    #[pyo3(text_signature = "($self, volume: float)")]
    pub fn set_volume(&self, volume: f64) -> NeosynthResult<()> {
        self.0.player.set_volume(volume)
    }
    /// Get the current speaking rate
    #[pyo3(text_signature = "($self) -> float")]
    pub fn get_rate(&self) -> NeosynthResult<f64> {
        if !Self::is_prosody_supported()? {
            Ok(-1.0)
        } else {
            Ok(self.0.synthesizer.Options()?.SpeakingRate()? / 0.06)
        }
    }
    /// Set the current speaking rate
    #[pyo3(text_signature = "($self, rate: float)")]
    pub fn set_rate(&self, value: f64) -> NeosynthResult<()> {
        if Self::is_prosody_supported()? {
            Ok(self
                .0
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
    /// Get the voice pitch
    #[pyo3(text_signature = "($self) -> float")]
    pub fn get_pitch(&self) -> NeosynthResult<f64> {
        Ok(self.0.synthesizer.Options()?.AudioPitch()? * 50.0)
    }
    /// Set the voice pitch
    #[pyo3(text_signature = "($self, pitch: float)")]
    pub fn set_pitch(&self, value: f64) -> NeosynthResult<()> {
        Ok(self.0.synthesizer.Options()?.SetAudioPitch(value / 50.0)?)
    }
    /// Get the current voice
    #[pyo3(text_signature = "($self) -> neosynth.VoiceInfo")]
    pub fn get_voice(&self) -> NeosynthResult<VoiceInfo> {
        Ok(self.0.synthesizer.Voice()?.into())
    }
    /// Set the current voice
    #[pyo3(text_signature = "($self, voice: neosynth.VoiceInfo)")]
    pub fn set_voice(&self, voice: &VoiceInfo) -> NeosynthResult<()> {
        Ok(self
            .0
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
        self.0.speak(utterance.0)?;
        self.0.set_state(SynthState::Busy)?;
        self.0.player.play()?;
        Ok(())
    }
    /// Pause the speech
    #[pyo3(text_signature = "($self)")]
    pub fn pause(&self) -> NeosynthResult<()> {
        self.0.set_state(SynthState::Paused)?;
        self.0.player.pause()?;
        Ok(())
    }
    /// Resume the speech
    #[pyo3(text_signature = "($self)")]
    pub fn resume(&self) -> NeosynthResult<()> {
        self.0.set_state(SynthState::Busy)?;
        self.0.player.resume()?;
        Ok(())
    }
    /// Stop the speech
    #[pyo3(text_signature = "($self)")]
    pub fn stop(&self) -> NeosynthResult<()> {
        self.0.player.stop()?;
        self.0.clear_speech_queue()?;
        self.0.process_queue()?;
        Ok(())
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
