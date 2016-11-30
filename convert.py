import numpy
import cv2
import xlwings as xw
 
import sys, shutil, os.path

def main():
    os.chdir("FULL PATH TO THIS FOLDER")
	videoName = ''
	f = open('input.txt', 'r')
	fileName = f.readline()
	videoName = fileName
	
	if os.path.exists('frames'):
		shutil.rmtree('frames')		#removes frame folder
	os.mkdir('frames')
	
	print cv2.__version__		#check cv2 version
	vidcap = cv2.VideoCapture(videoName)
	success,image = vidcap.read()
	count = 0
	success = True
	while success:
	  success,image = vidcap.read()
	  print 'Read a new frame: ', success
	  cv2.imwrite("frames\\" + str(count) + ".jpg", image)     # save frame as JPEG file
	  count += 1

        
if __name__ == '__main__':
	main()
