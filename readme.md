# Introduction

This program takes the URL of a youtube video, then renders it onto excel so a pixelated video can be played

# Work Flow

1. The convertion of the video is done trough `MMEPEG` Library using Python, which is called inside the vba script

2. The python script extracts all the frames from the video and save them into a folder

3. The vba script reads the frames and color cells based on the frames color. The next frame is drawn below the previous frame. 

4. After all frames are rendered, the play button scrolls through the rendered cells like a flip book.

