from pydub import AudioSegment
from pydub.playback import play
 
def change_playback_speed(sound, speed=1.0):
    # Manipulate the frame rate to change the speed
    sound_with_altered_frame_rate = sound._spawn(sound.raw_data, overrides={
         "frame_rate": int(sound.frame_rate * speed)
      })
    # Return sound with standard frame rate to maintain correct pitch
    return sound_with_altered_frame_rate.set_frame_rate(sound.frame_rate)
 
# Load an existing audio file
sound = AudioSegment.from_file(r"C:\Users\james680384\Downloads\Documents Completed.mp3")
 
# Change the speed (e.g., 1.5x speed)
speed_changed_sound = change_playback_speed(sound, 0.8)
 
# Export the modified sound to a new file
speed_changed_sound.export("output_sound.mp3", format="mp3")
 
# Optionally play the sound
play(speed_changed_sound)