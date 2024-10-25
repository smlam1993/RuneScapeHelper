import requests
import os
import json
import xlsxwriter
from bs4 import BeautifulSoup

os.system('cls')

SOUP_PARSER = 'html.parser'
base_url = 'https://runescape.wiki'

def get_quest_series():
    quest_series_page = requests.get(base_url + '/w/List_of_quest_series')
    quest_series_page_soup = BeautifulSoup(quest_series_page.content, SOUP_PARSER)

    # for tag in quest_series_page_soup.find_all(attrs={'class' : 'mw-header'}):
    headers = quest_series_page_soup.find_all('h3')
    series_count = 37

    quest_series = {}

    for i in range(series_count):
        header = headers[i]
        series_name = header.find('span', attrs = {'class': 'mw-headline'}).text
        table = header.find_next_sibling('table')
        rows = table.css.select('td:nth-child(2)')
        quests = []
        for row in rows:
            quest_name = row.find('a').text
            quests.append(quest_name)
            
        quest_series[series_name] = quests
        
    return quest_series

def get_runescape_data():
    with open('runescapeData_fixed.json', 'r') as jsonFile:
        runescape_data = json.load(jsonFile)

    runescape_dict = {}
    for obj in runescape_data:
        runescape_dict[obj['name']] = obj
    
    return runescape_dict

def parse_items(item_reqs):
    items_list = []
    for item_req in item_reqs:
        count = item_req['count']
        name = item_req['name']
        if(name.lower() == 'none'.lower() or count == 0):
            continue
        
        items_list.append(f"{count} {name}")
        
    return items_list

def parse_quests(quest_reqs):
    return quest_reqs

def parse_skills(skill_reqs):
    skills = []
    for skill_req in skill_reqs:
        level = skill_req['level']
        name = skill_req['name']
        if(name.lower() == 'none'.lower()):
            continue
        
        skills.append(f"{level} {name}")
        
    return skills

def generate_spreadsheet(quest_series, runescape_dict):
    workbook = xlsxwriter.Workbook('test.xlsx')
    
    quest_header_format = workbook.add_format()
    quest_header_format.set_bold()
    quest_header_format.set_pattern(1)
    quest_header_format.set_font_size(16)
    quest_header_format.set_bg_color('yellow')
    quest_header_format.set_align('center')
    quest_header_format.set_bottom(6)
    
    reqs_header_format = workbook.add_format()
    reqs_header_format.set_bold()
    reqs_header_format.set_pattern(1)
    quest_header_format.set_font_size(14)
    reqs_header_format.set_bg_color('green')
    reqs_header_format.set_align('center')
    reqs_header_format.set_border(1)
    
    data_cell_format = workbook.add_format()
    data_cell_format.set_pattern(1)
    data_cell_format.set_bg_color('#D3D3D3')
    data_cell_format.set_right(1)
    
    for series_name in quest_series:
        quest_names = quest_series[series_name]
        
        MAX_SKILL_COUNT = 0
        MAX_ITEM_COUNT = 0
        MAX_QUEST_COUNT = 0
        
        organized_quest_series = {} # quest - skills, pre-quests, items
        
        for quest_name in quest_names:
            if(quest_name not in runescape_dict):
                continue
            
            quest_info = runescape_dict[quest_name]            
            skills = parse_skills(quest_info['skill_reqs'])
            items = parse_items(quest_info['item_reqs'])
            parsed_quests = parse_quests(quest_info['quest_reqs'])
            
            if(len(skills) > MAX_SKILL_COUNT):
                MAX_SKILL_COUNT = len(skills)
                
            if(len(parsed_quests) > MAX_QUEST_COUNT):
                MAX_QUEST_COUNT = len(parsed_quests)
                
            if(len(items) > MAX_ITEM_COUNT):
                MAX_ITEM_COUNT = len(items)
                
            organized_quest_series[quest_name] = {
                'skills' : skills,
                'items' : items,
                'quests' : parsed_quests,
            }
        
        worksheet = workbook.add_worksheet(series_name)
        current_column = 0
        for parsed_quest in organized_quest_series:
            worksheet.set_column_pixels(current_column, current_column, 10 * len(parsed_quest), data_cell_format)
            worksheet.write(0, current_column, parsed_quest, quest_header_format)
            current_row = 1
            
            worksheet.write(current_row, current_column, 'Skill Levels', reqs_header_format)
            current_row += 1
            for i in range(MAX_SKILL_COUNT):
                skills = organized_quest_series[parsed_quest]['skills']
                if(len(skills) <= i):
                    break
                
                value = skills[i]
                worksheet.write(current_row + i, current_column, value, data_cell_format)
            current_row += MAX_SKILL_COUNT + 1
            
            worksheet.write(current_row, current_column, 'Items', reqs_header_format)
            current_row += 1
            for i in range(MAX_ITEM_COUNT):
                items = organized_quest_series[parsed_quest]['items']
                if(len(items) <= i):
                    break
                
                value = items[i]
                worksheet.write(current_row + i, current_column, value, data_cell_format)
            current_row += MAX_ITEM_COUNT + 1
            
            worksheet.write(current_row, current_column, 'Quests', reqs_header_format)
            current_row += 1
            for i in range(MAX_QUEST_COUNT):
                quests = organized_quest_series[parsed_quest]['quests']
                if(len(quests) <= i):
                    break
                
                value = quests[i]
                worksheet.write(current_row + i, current_column, value, data_cell_format)
            current_row += MAX_QUEST_COUNT + 1
            
            current_column += 1
    
    workbook.close()

def generate_item_spreadsheet(runescape_data):
    workbook = xlsxwriter.Workbook('items.xlsx')
    
    items = {}
    
    for x in runescape_data:
        for item in runescape_data[x]['item_reqs']:
            count = item['count']
            name = item['name']
            if name not in items:
                items[name] = 0
            items[name] += count
    
    column_index = 0
    row_index = 0
    worksheet = workbook.add_worksheet('Items')
    worksheet.write(row_index,column_index, "Name")
    worksheet.write(row_index,column_index + 1, "Count")
    
    for item in items:
        row_index += 1
        worksheet.write(row_index, column_index, item)
        worksheet.write(row_index, column_index + 1, items[item])
    
    worksheet.autofit()
    workbook.close()

runescape_data = get_runescape_data()
generate_spreadsheet(get_quest_series(), runescape_data)
generate_item_spreadsheet(runescape_data)