import subprocess
import pandas
import sys

def generate_cues(data_frame, identifying_character, midi_patch):
    cue_number = 0

    first_column = get_first_numeric_column(data_frame)
    total_cues = len(data_frame.columns[first_column:])

    print(f"Creating {total_cues} cues...")

    for cue in data_frame.columns[first_column:]:
        channels_to_unmute, channels_to_mute = get_channel_mute_data(data_frame, cue, identifying_character)
        
        create_cue(("Q" + str(cue)), str(cue), channels_to_unmute, str(len(channels_to_unmute) + len(channels_to_mute)))

        percent = round(((cue_number + 1) / total_cues) * 100)
        sys.stdout.write("\033[K") 
        print(f"'Q{cue}' created. ({percent}%)", end="\r", flush=True)

        cue_number += 1

    print("\nCreated {total_cues} cues.")

def get_channel_mute_data(data_frame, cue, identifying_character):
    channels_to_unmute = []
    channels_to_mute = []
    for _, row in data_frame.iterrows():
        if pandas.notna(row[cue]) and row[cue] == identifying_character:
            mic_num = int(row.iloc[0])
            channels_to_unmute.append(mic_num)
        else:
            mic_num = int(row.iloc[0])
            channels_to_mute.append(mic_num)
    return channels_to_unmute, channels_to_mute

def get_first_numeric_column(data_frame):
    for column in data_frame.columns:
        if isinstance(column, (int, float)):
            return data_frame.columns.get_loc(column)
    return None

def create_cue(cue_name, cue_number, channels_to_unmute, max_channels, midi_patch=1):
    channel_unmute_string = "{" + ", ".join(str(channel) for channel in channels_to_unmute) + "}"
    stdout, stderr = run_apple_script(
        apple_script
            .replace("{MAX_CHANNELS}", str(max_channels))
            .replace("{CHANNELS_TO_UNMUTE}", channel_unmute_string)
            .replace("{CUE_NAME}", str(cue_name))
            .replace("{CUE_NUMBER}", str(cue_number))
            .replace("{MIDI_PATCH}", str(midi_patch))
        )
    
    if stderr:
        print(stderr.decode("utf-8"))

def run_apple_script(script):
    process = subprocess.Popen(['osascript', '-e', script], stdout=subprocess.PIPE, stderr=subprocess.PIPE)
    stdout, stderr = process.communicate()
    return stdout, stderr

apple_script = """
tell application "QLab"
    set midiDevice to 1
    set channelsToUnmute to {CHANNELS_TO_UNMUTE}
    set totalChannels to {MAX_CHANNELS}

    make front workspace type "Group"
    set groupCue to last item of (selected of front workspace as list)

    tell groupCue
        set q name to "{CUE_NAME}"
        set q number to "{CUE_NUMBER}"
    end tell

    repeat with ch from 1 to totalChannels
        set ccNumber to (ch - 1)

        if ch is in channelsToUnmute then
            set ccValue to 127
        else
            set ccValue to 0
        end if

        make front workspace type "MIDI"
        set targetCue to last item of (selected of front workspace as list)
        set newCueID to uniqueID of targetCue

        if targetCue is not missing value then
            tell targetCue
                -- set patch to {MIDI_PATCH}
                set command to control_change
                set channel to 1
                set byte one to ccNumber
                set byte two to ccValue
                set q number to ""
                if ch is in channelsToUnmute then
                    set q name to "Channel " & ch & " Unmute"
                else
                    set q name to "Channel " & ch & " Mute"
                end if
            end tell

            move targetCue to end of groupCue
        else
            display dialog "Failed to create MIDI cue for channel " & ch
        end if
    end repeat
end tell
"""
