import requests
import os
import json
import re
from bs4 import BeautifulSoup


# classes
class QuestRequirement:
    def __init__(self, name, href):
        self.name = name
        self.href = href
        self.item_reqs = []
        self.recommendations = []
        self.quest_reqs = []
        self.skill_reqs = []

    def toJSON(self):
        return json.dumps(self.__dict__, cls=ComplexEncoder, indent=4)


class Skill:
    def __init__(self, level, name):
        self.level = level
        self.name = name

    def toJSON(self):
        return json.loads(json.dumps(self.__dict__, cls=ComplexEncoder, indent=4))


class Item:
    def __init__(self, count, name):
        self.count = count
        self.name = name

    def toJSON(self):
        return json.loads(json.dumps(self.__dict__, cls=ComplexEncoder, indent=4))


class ComplexEncoder(json.JSONEncoder):
    def default(self, obj):
        if hasattr(obj, "toJSON"):
            return obj.toJSON()
        else:
            return json.JSONEncoder.default(self, obj)


# get quests/skills requirements


def process_quest_req(node):
    quests = []

    for x in node.css.iselect("td > ul > li > ul > li"):
        ignore_next = False
        for y in x.contents:
            if y.name == "ul" or ignore_next:
                ignore_next = False
                continue

            requirement = y.text
            if not hasattr(y, "contents"):
                requirement = y.parent.text
                ignore_next = True

            if requirement == "None":
                continue

            quests.append(requirement.strip())

    return quests


def process_skill_req(node):
    skills = []

    for listItem in node.find_all("li"):
        split_values = [z for z in listItem.text.split(" ", 2) if z]
        try:
            level = int(split_values[0])
            name = split_values[1]
            skills.append(Skill(level, name))
        except ValueError:
            skills.append(Skill(0, listItem.text))

    return skills


def parse_quest_skill_requirements(questRequirement, quest_overview):
    quest_skill_requirements = quest_overview.find(
        string="Requirements"
    ).parent.nextSibling

    for content in quest_skill_requirements.contents:
        if content == "\n":
            continue

        if "Quests" in content.text:
            questRequirement.quest_reqs = process_quest_req(content)
            continue

        if "ul" == content.name:
            questRequirement.skill_reqs += process_skill_req(content)
            continue

        if "Ironmen" in content.text:
            questRequirement.skill_reqs += process_skill_req(content.find("ul"))
            continue


# get item requirements
def process_item_req(node):
    count = 1
    name = node.text

    anchor_tag = node.find("a")
    if anchor_tag is not None:
        name = node.find("a").get("title")

    matches = re.findall(r"[\d]+[,]?[\d]*", node.contents[0].text)
    if matches:
        count = int(matches[0].replace(",", "").strip())

    return Item(count, name)


def parse_item_requirements(questRequirement, quest_overview):
    items_required = quest_overview.find(string="Items required").parent.nextSibling
    all_lists = items_required.find_all("ul")
    for x in all_lists:
        for y in x.contents:
            if "\n" == y or y.text == "None":
                continue

            if y.name == "li":
                item = process_item_req(y)
                questRequirement.item_reqs.append(item)


# recommended
def parse_recommendations(questRequirement, quest_overview):
    recommendation_row = quest_overview.find(string="Recommended")
    if recommendation_row is None:
        return

    recommendations = recommendation_row.parent.nextSibling.find("ul").contents
    for x in recommendations:
        if "\n" == x:
            continue

        questRequirement.recommendations.append(x.text.replace("  ", " ").strip())


# start
os.system("cls")

SOUP_PARSER = "html.parser"
base_url = "https://runescape.wiki"

quests_page = requests.get(base_url + "/w/List_of_quests")
quests_page_soup = BeautifulSoup(quests_page.content, SOUP_PARSER)

quests = []
for tag in quests_page_soup.css.select("table.wikitable td:first-child a"):
    quest_name = tag.text
    href = tag.get("href")
    print("Processing " + quest_name)

    quest_request = requests.get(base_url + "/" + href + "/Quick_guide")
    quest_soup = BeautifulSoup(quest_request.content, SOUP_PARSER)
    quest_overview = quest_soup.find(id="Overview").parent.find_next_sibling("table")

    questRequirement = QuestRequirement(quest_name, href)
    parse_quest_skill_requirements(questRequirement, quest_overview)
    parse_item_requirements(questRequirement, quest_overview)
    parse_recommendations(questRequirement, quest_overview)

    quests.append(questRequirement)


def obj_dict(obj):
    return obj.__dict__


with open("runescapeData.json", "w") as runescapeDataFile:
    runescapeDataFile.write(
        json.dumps(
            quests, default=obj_dict, cls=ComplexEncoder, indent=4, sort_keys=True
        )
    )
