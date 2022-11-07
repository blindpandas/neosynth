# coding: utf-8

"""Shows the usage of Neosynth."""


import logging
import neosynth


FORMAT = '%(levelname)s %(name)s %(asctime)-15s %(filename)s:%(lineno)d %(message)s'
logging.basicConfig(format=FORMAT)
SSML = """
<speak version="1.0" xmlns="http://www.w3.org/2001/10/synthesis" xml:lang="en">
<s>Hello there!</s>
<mark name="mark1"/>
<p>Here comes a scilence</p>
<break time="1500ms"/>
<s>Goodbye!</s>
</speak>
""".strip()

class EventSink:
    """
    Required to handle synthesizer events.
    Must have the following two methods.
    """

    def on_state_changed(self, new_state):
        print(f"New state is: {new_state}")

    def on_bookmark_reached(self, bookmark):
        print(f"Bookmark reached: {bookmark}")

    def log(self, message, level):
        print(f"LOG {level}: {message}")


def main():
    # Setup the synthesizer
    synth = neosynth.Neosynth(EventSink())
    synth.set_pitch(50.0)
    synth.set_rate(30.0)
    synth.set_volume(75.0)
    synth.speak_text("Hello")
    synth.speak_ssml(SSML)

if __name__ == '__main__':
    main()