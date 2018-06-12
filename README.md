# Disclaimer
This code is intended to be used only for the CSCI 3104 class at CU for uploading feedback automatically from Google Sheet to Moodle.

# How to use it
1. **student_info_parser.py** - This code is used to build a Email to Moodle ID mapping and save it on disk as a python pickle file
This should be run only to build the mapping once (unless the mapping changes)
2. For above to work you need to modify the code according to html (moodle page with the list of students which has their email and moodle id both) you are scraping (Moodle tends to keep on changing the UI)
3. Replace dummy credentials in **moodle_credentials.json** and **client_secret.json** (Google Sheet API credentials)
4. **do_not_upload_for.txt** has email ids for which you don't want to upload feedback
5. **delete_feedback_for.txt** has email ids for which you want to delete uploaded feedback.
6. **sheet_to_moodle.py** has instructions and I would add more instructions in future.


**Note** : I hastily put it and it might be tough to understand some parts.
       I will be around to help :).
       You can email me at my email addresses - firstname.lastname@gmail.com /
                                                firstname.lastname@colorado.edu