A simple python application that detects changes on a specified rectangle of your screen. 
If a change is happening (for example a new slide of a presentation), the programm will take a screenshot. 
The Output is a Presentation with all screenshots taken in the given time frame.

The algorithm will take a screenshit every half a second and has a waiting time to change the window to the video of 30 seconds.
It also asks for 2 parameters, the maximum slide numbers (to prevent overcrowding your directory) and a time limit to know when the video is over.

To run the programm there has to be a folder called "Slides" in the same directory as the programm. 
The output will be a presentation called "AllSlides.pptx" with, well, all slides.

I uploaded the .py and .ipynb they should be identical so feel free to use either.
