1. use pyside6 GUI to create the application
2. when user select the folder, the display will show all the file name (Excel, CSV, txt file)
3. the application have a textbox for user to key in the keyword to search in the file
4. the keyword can be multiple keyword with semicolon (example : product, branch)
5. when found the specific keyword in that file the display box will show the file name that contain that keyword
6. the shown file allowed user to click and direct open from the display box.
7. have a checkbox that allowed the application to match exactly the keyword or contain the keyword ( tick : matched, untick : just_contain_keyword), default is matched
8. use Qsetting to save the folder path and setting and use it during application initialize and closing the application.
9. for file path using pathlib