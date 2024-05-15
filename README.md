# Video-Mark-Automation-Script
COMP 467 -  Project 3: "The Crucible" 

# Project Description: 
The script aimed to streamline virtual production workflows through automation. It automated the conversion of file paths, optimized output file formats, and implemented a robust backup strategy for critical project data. Additionally, it enhanced output files with thumbnails for visual cues and automated the upload process to FrameIO for improved collaboration and accessibility. This initiative revolutionizes virtual production workflows, empowering teams with efficiency and focus to excel in their creative pursuits.

1. Reuse Proj 1​
2. Add argparse to input baselight file (--baselight),  xytech (--xytech) from proj 1​
3. Populate new database with 2 collections: One for Baselight (Folder/Frames) and Xytech (Workorder/Location)​
4. Download my amazing VP video, https://mycsun.box.com/s/v55rwqlu5ufuc8l510r8nni0dzq5qki7Links to an external site.​
5. Run script with new argparse command --process <video file>  ​
6. From (5) Call the populated database from (3), find all ranges only that fall in the length of video from (4)
7. Using ffmpeg or 3rd party tool of your choice, to extract timecode from video and write your own timecode method to convert marks to timecode​
8. New argparse--output parameter for XLS with flag from (5) should export same CSV export as proj 1 (matching xytech/baselight locations), but in XLS with new column from files found from (6) and export their timecode ranges as well​
9. Create Thumbnail (96x74) from each entry in (6), but middle most frame or closest to. Add to XLS file to it's corresponding range in new column ​
10. Render out each shot from (6) using (7) and upload them using API to frame.io (https://developer.frame.io/api/reference/)

---------------------------------------------------------------------------------

# Deliverables​

1. Copy/Paste code​
2. Excel file with new columns noted on Solve (8) and (9)​
3. Screenshot of Frame.io account (10)
