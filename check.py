import os
from mutagen.flac import FLAC
from mutagen.mp3 import EasyMP3

def check_audio_files(folder):
    title_is_kuwo = []
    no_lyrics = []
    for fname in os.listdir(folder):
        if fname.lower().endswith('.flac'):
            try:
                audio = FLAC(os.path.join(folder, fname))
                title = audio.get('title', [''])[0]
                lyrics = audio.get('LYRICS', ['']) or audio.get('lyrics', [''])
                if title == 'kuwo':
                    title_is_kuwo.append(fname)
                if not lyrics or lyrics[0].strip() == '':
                    no_lyrics.append(fname)
            except Exception:
                continue
        elif fname.lower().endswith('.mp3'):
            try:
                audio = EasyMP3(os.path.join(folder, fname))
                title = audio.get('title', [''])[0]
                lyrics = audio.get('lyrics', [''])
                if title == 'kuwo':
                    title_is_kuwo.append(fname)
                if not lyrics or lyrics[0].strip() == '':
                    no_lyrics.append(fname)
            except Exception:
                continue
    return title_is_kuwo, no_lyrics

def print_results(title_is_kuwo, no_lyrics):
    print("\n" + "="*50)
    print("标题为kuwo的文件 (总数: {})".format(len(title_is_kuwo)))
    print("="*50)
    for f in title_is_kuwo:
        print(f)
    print("\n" + "="*50)
    print("不含歌词的文件 (总数: {})".format(len(no_lyrics)))
    print("="*50)
    for f in no_lyrics:
        print(f)

def reset_kuwo_titles(folder, kuwo_files):
    for fname in kuwo_files:
        filepath = os.path.join(folder, fname)
        new_title = os.path.splitext(fname)[0]
        if fname.lower().endswith('.flac'):
            audio = FLAC(filepath)
            audio['title'] = new_title
            audio.save()
        elif fname.lower().endswith('.mp3'):
            audio = EasyMP3(filepath)
            audio['title'] = new_title
            audio.save()
        print(f"已重设 {fname} 的标题为 {new_title}")

if __name__ == "__main__":
    folder = os.getcwd()
    title_is_kuwo, no_lyrics = check_audio_files(folder)
    print_results(title_is_kuwo, no_lyrics)
    if title_is_kuwo:
        confirm = input(f"发现 {len(title_is_kuwo)} 个标题为kuwo的文件，是否重设标题为文件名？(y/n): ")
        if confirm.lower() == 'y':
            reset_kuwo_titles(folder, title_is_kuwo)
        else:
            print("取消重设。")