# Google-Photos-Metadata-Fixer (Free)
This project is a replica of Google Photo Metadata Fixer, a tool that helps users to fix their photo metadata. The tool allows users to bulk modify their photos’ date and time properties.

Based on https://github.com/ikanwalkhalsa/google-photos-metadata-fixer/

This replica is written in Python. The tool provides a command-line interface (CLI) that just requires the location of zip files and handle the rest from there.

Features:
* Modify photos’ date and time in bulk.
* exif pictures with goelocation and people
* unzip, or use already flat unzipped takeout files
* Provide a CLI for easy interaction with the tool.
* Handle errors and exceptions gracefully to ensure smooth execution.

Installation:
<br/>To use the script, you just need Python installed on your system. And then follow the given steps.
* Download the script 'google_photos_metadata_fixer.py' on your system
* run 'pip install piexif Pillow pywin32'
* run 'python google_photos_metadata_fixer.py "PATH OF FOLDER CONTAINTING ZIP FILES OR FLAT PICTURES AND JSONS"'

Note:
<br/>Here are a few things that can help you in the process.
* Make sure your zip files follow the same naming convention as google (takeout-*.zip)
* Source Folder Location can be modified on line 9
* Output Folder Location can be modified on line 10

How it Works:
* Scans all the takeout zips in the given folder
* Unzip the files in the same folder location
* Merges folders from different zips into a single intermediate Folder
* Looks and pairs JSON metadata for all file types
* Moves all files to output destination and updates the metadata
* If no JSON found for the file the file will be moved to /FAILED dir
