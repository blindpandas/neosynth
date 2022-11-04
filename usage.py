# coding: utf-8

"""Shows the usage of Neosynth."""


import logging
import neosynth


FORMAT = '%(levelname)s %(name)s %(asctime)-15s %(filename)s:%(lineno)d %(message)s'
logging.basicConfig(format=FORMAT)


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
    # create the speech utterance
    ut = neosynth.SpeechUtterance()
    ut.add_text("Hello there.")
    ut.add_text("And another thing.")
    ut.add_text("Goodbye!")
    ut.add_ssml("""
        <speak version="1.0" xml:lang="en-US">
          <s>Hello Musharraf</s>
          <mark name="mark1"/>
          <s>Rust is the best</s>
          <mark name="mark2"/>
          <s>Bye</s>
          <mark name="it_works"/>
        </speak>
    """.strip())
    # Speak 
    synth.speak(ut)


if __name__ == '__main__':
    main()