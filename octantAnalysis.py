from datetime import datetime
start_time = datetime.now()

import streamlit as st
import os
os.system("cls") # clearing screen
os.chdir(r'C:\Users\Aman Jha\Documents\GitHub\2001EE22_2022\proj2') # making the parent directory as current directory

def proj_octant_gui():
	
	# Find more emojis here: https://www.webfx.com/tools/emoji-cheat-sheet/
	st.set_page_config(page_title="CS384 Project", page_icon=":dizzy:",layout="wide") # basic page configuration
	
	st.sidebar.success("Select a page from above for further action:")
	with st.container():
		st.title("Project 2 of Python!!! :fire:")
		st.header("Made by Aman Jha and Anuradha Das Group :star2:")
		st.subheader("2001EE22 and 2001CB10")
		st.write('''
		This is a multipage website. Currently you are on the home page. :smile:''')
		st.write("---")

proj_octant_gui()

from platform import python_version
ver = python_version()

if ver == "3.8.10":
	print("Correct Version Installed")
else:
	print("Please install 3.8.10. Instruction are present in the GitHub Repo/Webmail. Url: https://pastebin.com/nvibxmjw")

#This shall be the last lines of the code.
end_time = datetime.now()
print('Duration of Program Execution: {}'.format(end_time - start_time))