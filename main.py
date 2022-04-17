import os


# def makingPPT():
#     from pptx import Presentation
#     from pptx.util import Inches
#     presentation = Presentation()
#
#     presentation.slide_height = Inches(9)
#     presentation.slide_width = Inches(16)
#
#     slide_layout = presentation.slide_layouts[6]
#
#     slide = presentation.slides.add_slide(slide_layout)
#
#     left = Inches(0)
#     top = Inches(0)
#     width = presentation.slide_width
#     height = presentation.slide_height
#
#     pic = slide.shapes.add_picture('슬라이드1.jpg', left, top, width, height)
#
#     slide2 = presentation.slides.add_slide(slide_layout)
#     pic2 = slide2.shapes.add_picture('슬라이드2.jpg', left, top, width, height)
#
#     presentation.save('demo.pptx')


def text_from_word(path, doc_name):
    import docx2txt
    import time
    import pyautogui

    doc_path = os.path.join(path, doc_name)
    result = docx2txt.process(doc_path)

    with open(f"{doc_path}.txt", "w", encoding='utf8') as f:
        f.write(result)

    pyautogui.click(pyautogui.locateOnScreen("../resource/chrome.png", confidence=0.7))
    time.sleep(2)

    time.sleep(1)
    pyautogui.keyDown('ctrl')
    pyautogui.press('t')
    pyautogui.keyUp('ctrl')

    time.sleep(1)
    pyautogui.write('https://app.typecast.ai/ko/login')

    time.sleep(1)
    pyautogui.press('enter')

    time.sleep(10)
    a = pyautogui.click(pyautogui.locateOnScreen("../resource/typecast_login.png", confidence=0.7))
    # pyautogui.press('enter')

    print(a)

    # time.sleep(5)
    # pyautogui.keyDown('shift')
    # pyautogui.press('tab')
    #
    # time.sleep(1)
    #
    # pyautogui.press('tab')
    # pyautogui.keyUp('shift')
    #
    # time.sleep(1)
    # pyautogui.press('enter')
    #
    # time.sleep(5)
    # pyautogui.press('tab')
    # time.sleep(1)
    # pyautogui.press('tab')

    time.sleep(1)
    pyautogui.press('enter')
    # pyautogui.click(pyautogui.locateOnScreen("../resource/andrew.png", confidence=0.7))

    time.sleep(10)
    pyautogui.press('f6')

    time.sleep(1)
    pyautogui.write('https://app.typecast.ai/ko/dashboard/my-project')

    time.sleep(1)
    pyautogui.press('enter')

    time.sleep(1)
    pyautogui.click(pyautogui.locateOnScreen("../resource/new_project.png", confidence=0.7))

    time.sleep(1)
    pyautogui.click(pyautogui.locateOnScreen("../resource/new_project_step2.png", confidence=0.7))

    time.sleep(3)
    pyautogui.click(pyautogui.locateOnScreen("../resource/new_project_step2.png", confidence=0.7))

    time.sleep(1)
    pyautogui.click(pyautogui.locateOnScreen("../resource/fill_content.png", confidence=0.7))

    time.sleep(1)
    pyautogui.write(result)

    time.sleep(1)
    pyautogui.keyDown('ctrl')
    pyautogui.press('a')
    pyautogui.keyUp('ctrl')

    time.sleep(1)
    pyautogui.click(pyautogui.locateOnScreen("../resource/emotion.png", confidence=0.7))

    time.sleep(1)
    pyautogui.click(pyautogui.locateOnScreen("../resource/download.png", confidence=0.7))

    time.sleep(1)
    pyautogui.click(pyautogui.locateOnScreen("../resource/no_title.png", confidence=0.7))

    time.sleep(1)
    pyautogui.click(pyautogui.locateOnScreen("../resource/download_step2.png", confidence=0.7))

    time.sleep(1)
    pyautogui.keyDown('ctrl')
    pyautogui.press('a')
    pyautogui.keyUp('ctrl')

    time.sleep(1)
    pyautogui.write(doc_name)

    # result2 = re.sub("니다", "니다\n", result)
    # result3 = re.sub("지만", "지만\n", result2)
    #
    # with open(f"이유.txt", "w", encoding='utf8') as f:
    #     f.write(result3)
    #     # for element in result:
    #     #     f.write(f"file '{element}' \n")
    #
    # print(result3)


def save_as_image_from_slide(path_, slide_name_):
    from pptx_tools import utils
    pptx_file = os.path.join(path_, f'{slide_name_}')
    png_folder = pptx_file  # TODO: BUG
    utils.save_pptx_as_png(png_folder, pptx_file, True)


def make_short_video_per_image(path_, slide_name_, ffmpeg_exe_path_, mp3_path_):
    import re

    pptx_file = os.path.join(path_, f'{slide_name_}')
    png_folder = re.sub('.pptx', '', pptx_file)

    jpg_folder = f'{png_folder}_jpg'
    if not os.path.isdir(jpg_folder):
        os.mkdir(jpg_folder)

    from PIL import Image
    png_list = os.listdir(png_folder)
    for png in png_list:
        png_img_path = os.path.join(png_folder, png)
        img = Image.open(png_img_path).convert('RGB')

        jpg_path = os.path.join(jpg_folder, re.sub('.PNG', '.JPG', png))
        img.save(jpg_path, quality=100)

    short_video_folder = f'{png_folder}_short'
    if not os.path.isdir(short_video_folder):
        os.mkdir(short_video_folder)

    from mutagen.mp3 import MP3
    audio = MP3(mp3_path_)
    audio_length = audio.info.length

    jpg_list = os.listdir(jpg_folder)
    duration = round(audio_length / len(jpg_list), 2)

    frame_rate = 25

    # ffmpeg complex filter
    cf = f"scale=-2:10*ih,zoompan=z='min(zoom+0.0015,1.5)':d={duration * frame_rate}" \
         f":x='iw/2-(iw/zoom/2)':y='ih/2-(ih/zoom/2)',scale=-2:1080"

    # def execute_ffmpeg(cmd_):
    #     os.system(cmd_)
    # thread_list = []
    # num_requests = 0
    # with ThreadPoolExecutor(max_workers=10) as executor:
    for jpg in jpg_list:
        in_ = os.path.join(jpg_folder, jpg)
        # in_ = re.sub(" ", "%20", in_)

        out_file_name_1 = re.sub('슬라이드', '', jpg)
        out_file_name_2 = re.sub('.JPG', '', out_file_name_1)
        out_file_name_3 = f'{out_file_name_2.zfill(2)}.MP4'
        out_ = os.path.join(short_video_folder, out_file_name_3)
        # out_ = re.sub(" ", "%20", out_)
        # ffmpeg_exe_path__ = re.sub(" ", "%20", ffmpeg_exe_path_)
        cmd = f'cmd /C ("{ffmpeg_exe_path_}" -y -loop 1 -r {frame_rate} -i "{in_}" ' \
              f'-filter_complex "{cf}" -shortest -c:v libx264 -t {duration} "{out_}")'

        print(cmd)
        os.system(cmd)

        #     future = executor.submit(execute_ffmpeg, cmd)
        #     thread_list.append(future)
        # for execution in concurrent.futures.as_completed(thread_list):
        #     execution.result()
        # thread = threading.Thread(target=execute_ffmpeg, args=cmd)
        # thread.start()


def merge_short_into_complete_video(path_, slide_name_, ffmpeg_exe_path_, mp3_path_):
    import re
    pptx_file = os.path.join(path_, f'{slide_name_}')
    png_folder = re.sub('.pptx', '', pptx_file)

    short_video_folder = f'{png_folder}_short'
    complete_video_path = f'{png_folder}.MP4'

    short_list = os.listdir(short_video_folder)
    short_list_paths = [os.path.join(short_video_folder, x) for x in short_list]

    with open(f"{png_folder}_list.txt", "w", encoding='utf8') as f:
        # f.write(f"file '{path_}\\..\\인트로1.mp4' \n")
        for element in short_list_paths:
            f.write(f"file '{element}' \n")
        # f.write(f"file '{path_}\\..\\아웃트로.mp4' \n")

    # from moviepy.editor import VideoFileClip
    # clip = VideoFileClip(f'{path_}\\..\\인트로1.mp4')
    # print(clip.duration)

    cmd = f'cmd /C ("{ffmpeg_exe_path_}" ' \
          f'-y -f concat -safe 0 -r 25 -i "{png_folder}_list.txt" -i "{mp3_path_}" ' \
          f' -c:v libx264 -c:a libmp3lame  "{complete_video_path}")'
    # -vf "scale=1920x1080"
    os.system(cmd)


def make_subtitle_using_brew(path_, slide_name_, ffmpeg_exe_path_, mp3_path_):
    import pyautogui
    import time

    pyautogui.click(pyautogui.locateOnScreen("../resource/brew.png", confidence=0.7))
    time.sleep(5)

    ###############################

    pyautogui.click(pyautogui.locateOnScreen("../resource/new_video.png", confidence=0.7))

    time.sleep(2)
    pyautogui.press('f4')

    time.sleep(2)
    pyautogui.keyDown('ctrl')
    pyautogui.press('a')
    pyautogui.keyUp('ctrl')

    time.sleep(1)
    pyautogui.typewrite(path_)

    time.sleep(1)
    pyautogui.press('enter')

    time.sleep(1)
    pyautogui.click(pyautogui.locateOnScreen("../resource/mp4.png", confidence=0.7))

    time.sleep(1)
    pyautogui.press('enter')
    # pyautogui.click(pyautogui.locateOnScreen("../resource/file_name.png", confidence=0.7))

    time.sleep(1)
    pyautogui.click(pyautogui.locateOnScreen("../resource/confirm.png", confidence=0.7))


def append_inout_video_and_complete_video(path_, slide_name_, ffmpeg_exe_path_):
    import re
    pptx_file = os.path.join(path_, f'{slide_name_}')
    png_folder = re.sub('.pptx', '', pptx_file)

    complete_video_path = f'{png_folder}.MP4'
    complete_inout_video_path = f'{png_folder}_inout.MP4'

    with open(f"{png_folder}_list_inout.txt", "w", encoding='utf8') as f:
        f.write(f"file '{path_}\\..\\인트로1.mp4' \n")
        f.write(f"file '{complete_video_path}' \n")
        f.write(f"file '{path_}\\..\\아웃트로.mp4' \n")

    cmd = f'cmd /C ("{ffmpeg_exe_path_}" ' \
          f'-y -f concat -safe 0 -r 25 -i "{png_folder}_list_inout.txt" ' \
          f' -acodec libmp3lame "{complete_inout_video_path}")'
    # -vf "scale=1920x1080"
    os.system(cmd)


if __name__ == '__main__':
    import sys

    args = sys.argv

    # if len(args) > 2:
    #     path = args[1]
    #     slide_name = args[2]
    # else:
    os.chdir(args[1])
    path = os.getcwd()
    file_list = os.listdir(path)
    pptx_list = [x for x in file_list if x.endswith('.pptx')]
    slide_name = pptx_list[0]
    print(slide_name)

    # if len(args) > 3:
    #     ffmpeg_exe_path = args[3]
    # else:
    parent_path = os.getcwd()
    ffmpeg_path = os.path.join(parent_path, '..', 'ffmpeg')
    ffmpeg_exe_path = os.path.join(ffmpeg_path, 'bin', 'ffmpeg.exe')

    # if len(args) > 4:
    #     mp3_name = args[4]
    # else:
    file_list = os.listdir(path)
    mp3_list = [x for x in file_list if x.endswith('.mp3')]
    mp3_name = mp3_list[0]
    mp3_path = os.path.join(path, mp3_name)
    print(mp3_path)

    # file_list = os.listdir(path)
    doc_list = [x for x in file_list if x.endswith('.docx')]
    doc_name = doc_list[0]

    # text_from_word(path, doc_name)

    # save_as_image_from_slide(path, slide_name)
    make_short_video_per_image(path, slide_name, ffmpeg_exe_path, mp3_path)
    merge_short_into_complete_video(path, slide_name, ffmpeg_exe_path, mp3_path)
    # append_inout_video_and_complete_video(path, slide_name, ffmpeg_exe_path)

    # print((path, slide_name, ffmpeg_exe_path, mp3_path))
    # make_subtitle_using_brew(path, slide_name, ffmpeg_exe_path, mp3_path)
