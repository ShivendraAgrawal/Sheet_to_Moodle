from pprint import pprint
import pickle
from bs4 import BeautifulSoup

# You may have to modify this file heavily
# The purpose of this file is to create a python dict with keys as
# email ids (must be lower case and taken from latest class roster) and
# values as the unique Moodle ID
# I used https://moodle.cs.colorado.edu/mod/assign/view.php?id=20689&action=grading
# page to save the html and scrap emails and Moodle id
soup = BeautifulSoup(open('student_info.html', encoding='utf-8').read(), 'html.parser')

students = []

for row in soup.find_all('tr'):
    moodle_id = row.find('input').get('name')[4:]
    name = row.find('a').text
    email = row.find_all('td')[2].text

    students.append({'moodle_id':moodle_id, 'name':name, 'email':email})

pprint(students)
print(len(students))

unique_names = set([row['name'] for row in students])
print(len(unique_names))

id_lookup = {}
for student in students:
    id_lookup[student['email'].lower()] = student['moodle_id']

with open('id_lookup.pickle', 'wb') as f:
    pickle.dump(id_lookup, f, pickle.HIGHEST_PROTOCOL)

# Just printing the dict to make sure it looks fine
with open('id_lookup.pickle', 'rb') as f:
    id_lookup = pickle.load(f)
pprint(id_lookup)
