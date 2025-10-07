import os
import sys
from mutagen.flac import FLAC
from mutagen.mp3 import MP3
from mutagen.id3._frames import TIT2, TPE1

def delete_lrc_files(folder):
    """Delete all .lrc files in the folder and subfolders."""
    for root, dirs, files in os.walk(folder):
        for file in files:
            if file.lower().endswith('.lrc'):
                filepath = os.path.join(root, file)
                try:
                    os.remove(filepath)
                    print(f"Deleted: {filepath}")
                except Exception as e:
                    print(f"Error deleting {filepath}: {e}")

def parse_filename(filename):
    """Parse filename to extract artist and title.
    If '&' exists, split on the first '-' after the last '&'.
    If no '&', split on the first '-'.
    """
    name = os.path.splitext(filename)[0]
    if '&' in name:
        # Find last '&'
        last_amp = name.rfind('&')
        # Part after last '&'
        after_amp = name[last_amp + 1:]
        if '-' in after_amp:
            # Find first '-' in after_amp
            first_dash = after_amp.find('-')
            # Artist: everything before last_amp + '&' + after_amp[:first_dash]
            artist = name[:last_amp + 1] + after_amp[:first_dash]
            title = after_amp[first_dash + 1:]
        else:
            # No '-' after last '&', treat as invalid
            artist = ''
            title = name
    else:
        # No '&', split on first '-'
        if '-' in name:
            first_dash = name.find('-')
            artist = name[:first_dash]
            title = name[first_dash + 1:]
        else:
            artist = ''
            title = name
    return artist.strip(), title.strip()

def clean_metadata(filepath, artist, title):
    """Clean and set metadata for flac or mp3 file."""
    ext = os.path.splitext(filepath)[1].lower()
    try:
        if ext == '.flac':
            audio = FLAC(filepath)
            audio.clear()
            audio['title'] = title
            if artist:
                audio['artist'] = artist
            audio.save()
        elif ext == '.mp3':
            audio = MP3(filepath)
            audio.delete()  # Clear all tags
            audio['TIT2'] = TIT2(encoding=3, text=title)  # Title with UTF-8 encoding
            if artist:
                audio['TPE1'] = TPE1(encoding=3, text=artist)  # Artist with UTF-8 encoding
            audio.save()
        print(f"Updated metadata for: {filepath}")
    except Exception as e:
        print(f"Error updating {filepath}: {e}")

def process_music_files(folder):
    """Process all flac and mp3 files in the folder and subfolders."""
    for root, dirs, files in os.walk(folder):
        for file in files:
            if file.lower().endswith(('.flac', '.mp3')):
                filepath = os.path.join(root, file)
                artist, title = parse_filename(file)
                clean_metadata(filepath, artist, title)

def main():
    if len(sys.argv) > 1:
        folder = sys.argv[1]
    else:
        folder = input("Enter the folder path to process: ").strip()
    
    if not os.path.isdir(folder):
        print("Invalid folder path.")
        return
    
    confirm = input("This will delete all .lrc files and modify metadata of flac/mp3 files. Confirm? (y/n): ").strip().lower()
    if confirm != 'y':
        print("Operation cancelled.")
        return
    
    print("Deleting .lrc files...")
    delete_lrc_files(folder)
    
    print("Processing music files...")
    process_music_files(folder)
    
    print("Done.")

if __name__ == "__main__":
    main()