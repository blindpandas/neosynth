# coding: utf-8

"""Shows the usage of Neosynth."""


import neosynth



class EventSink:
    """
    Required to handle synthesizer events.
    Must have the following two methods.
    """

    def on_state_changed(self, new_state):
        print(f"New state is: {new_state}")

    def on_bookmark_reached(self, bookmark):
        print(f"Bookmark reached: {bookmark}")



def main():
    # Setup the synthesizer
    synth = neosynth.Neosynth(EventSink())
    synth.set_pitch(50.0)
    synth.set_rate(30.0)
    synth.set_volume(75.0)
    # create the speech utterance
    ut = neosynth.SpeechUtterance()
    ut.add_text("Hello there.")
    ut.add_bookmark("bookmark1")
    ut.add_text("And another thing.")
    ut.add_bookmark("bookmark2")
    ut.add_text("Goodbye!")
    # Speak 
    synth.speak(ut)


if __name__ == '__main__':
    main()