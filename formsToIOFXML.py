'''
Python 3.9 or higher required

Download the entries as csv files from the Google Sheets and save them in the same folder as this script,
then run the script to generate the IOF XML files
'''
import xml.etree.ElementTree as ET
import csv
from datetime import datetime
import copy
import os

############### File paths ####################

INDIVIDUAL_PATH = 'NSM 2022 entries (Responses) - Individuella.csv'
RELAY_PATH = 'NSM 2022 entries (Responses) - Stafett.csv'

############### Main functions ################

def main():
    if os.path.isfile(INDIVIDUAL_PATH):
        with open(INDIVIDUAL_PATH, 'r', encoding='utf-8') as f:
            reader = csv.reader(f)
            next(reader) # skip header
            individual(reader)
    else:
        print('No file with individual entries found')

    if os.path.isfile(RELAY_PATH):
        with open(RELAY_PATH, 'r', encoding='utf-8') as f:
            reader = csv.reader(f)
            next(reader) # skip header
            relay(reader)
    else:
        print('No file with relay entries found')

        
def individual(reader):
    entry_lists = [entry_list(), entry_list()]
    class_ids = {'Men Elite': '1', 'Women Elite': '2', 'Party Animal': '3'}
    for row in reader:
        for competition in enumerate(row[5:]):
            if competition[1]:
                person_entry = ET.SubElement(entry_lists[competition[0]], 'PersonEntry')
                person = ET.SubElement(person_entry, 'Person')
                name = ET.SubElement(person, 'Name')
                family = ET.SubElement(name, 'Family')
                if len(row[2].split(' ')) > 1:
                    family.text = ' '.join(row[2].split(' ')[1:])
                given = ET.SubElement(name, 'Given')
                given.text = row[2].split(' ')[0]

                organisation = ET.SubElement(person_entry, 'Organisation')
                org_id = ET.SubElement(organisation, 'Id')
                org_id.text = get_club_id(row[3])
                org_name = ET.SubElement(organisation, 'Name')
                org_name.text = row[3]

                control_card = ET.SubElement(person_entry, 'ControlCard')
                control_card.text = row[4]

                comp_class = ET.SubElement(person_entry, 'Class')
                class_id = ET.SubElement(comp_class, 'Id')
                class_id.text = class_ids[competition[1]]
                class_name = ET.SubElement(comp_class, 'Name')
                class_name.text = competition[1]
                
                entry_time = ET.SubElement(person_entry, 'EntryTime')
                entry_time.text = datetime.strptime(row[0], '%d/%m/%Y %H:%M:%S').isoformat()

    for competition in enumerate(entry_lists):
        tree = ET.ElementTree(competition[1])
        ET.indent(tree, space="\t", level=0)

        comp_enums = {0: 'Sprint', 1: 'Long'}
        file_name = f'entries_{comp_enums[competition[0]].lower()}.xml'

        tree.write(file_name, encoding='utf-8', xml_declaration=True)


def relay(reader):
    relay_entry_list = entry_list()
    # Define classes
    class_ids = {'Men Elite': '1', 'Women Elite': '2', 'Mixed Elite': '3'}
    event = ET.SubElement(relay_entry_list, 'Event')
    for relay_class in class_ids:
        class_info = ET.SubElement(event, 'Class')
        class_id = ET.SubElement(class_info, 'Id')
        class_id.text = class_ids[relay_class]
        class_name = ET.SubElement(class_info, 'Name')
        class_name.text = relay_class
        for i in range(3):
            ET.SubElement(class_info, 'Leg')
    # Add teams
    team_id_counter = 1
    for row in reader:
        team_entry = ET.SubElement(relay_entry_list, 'TeamEntry')
        team_id = ET.SubElement(team_entry, 'Id')
        team_id.text = str(team_id_counter)
        team_id_counter += 1
        team_name = ET.SubElement(team_entry, 'Name')
        team_name.text = check_team_name(row[3])

        for i in range(3):
            team_entry_person = ET.SubElement(team_entry, 'TeamEntryPerson')
            person = ET.SubElement(team_entry_person, 'Person')

            name = ET.SubElement(person, 'Name')
            family = ET.SubElement(name, 'Family')
            name_index = 4 + i * 2
            if len(row[name_index].split(' ')) > 1:
                family.text = ' '.join(row[name_index].split(' ')[1:])
            given = ET.SubElement(name, 'Given')
            given.text = row[name_index].split(' ')[0]

            leg = ET.SubElement(team_entry_person, 'Leg')
            leg.text = str(i + 1)

            control_card = ET.SubElement(team_entry_person, 'ControlCard')
            control_card.text = row[name_index + 1]

        comp_class = ET.SubElement(team_entry, 'Class')
        class_id = ET.SubElement(comp_class, 'Id')
        class_id.text = class_ids[row[2]]
        class_name = ET.SubElement(comp_class, 'Name')
        class_name.text = row[2]

        entry_time = ET.SubElement(team_entry, 'EntryTime')
        entry_time.text = datetime.strptime(row[0], '%d/%m/%Y %H:%M:%S').isoformat()

    tree = ET.ElementTree(relay_entry_list)
    ET.indent(tree, space="\t", level=0)

    tree.write('entries_relay.xml', encoding='utf-8', xml_declaration=True)

################### Helper functions ###################

team_names = dict()
def check_team_name(team_name):
    if team_name.lower() in team_names:
        team_names[team_name.lower()] += 1
        team_name += ' ' + str(team_names[team_name.lower()])
    else:
        team_names[team_name.lower()] = 1
    return team_name

club_ids = dict()
def get_club_id(club_name):
    club_name = club_name.strip()
    if club_name not in club_ids:
        club_ids[club_name] = str(len(club_ids) + 1)
    return club_ids[club_name]
    
def entry_list():
    entry_list = ET.Element('EntryList')
    entry_list.set('xmlns', 'http://www.orienteering.org/datastandard/3.0')
    entry_list.set('xmlns:xsi', 'http://www.w3.org/2001/XMLSchema-instance')
    entry_list.set('iofVersion', '3.0')
    entry_list.set('createTime', datetime.now().isoformat())
    return copy.deepcopy(entry_list)


if __name__ == '__main__':
    main()