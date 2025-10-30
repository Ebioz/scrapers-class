import openpyxl
from openpyxl import Workbook
import re
import time

            # ---------- Create output excel file and define titles ----------
wb = Workbook()
ws = wb.active
abstractUrl = "https://manage.ercongressi.it/storage/ercongressi/article/pdf/254/14496-SIE_2025.pdf"
role0 = "Abstract author"
email = ""
session_description = ""

            # ---------- Create table title ----------
table_title = [
    'Name (incl. titles)',
    'Affiliation/Organisation and location',
    'Role',
    'Email',
    'Session Name',
    'Session Description',
    'Presentation Title',
    'Presentation Abstract',
    'Abstract URL',
    'Video URL',
]
ws.append(table_title)

            # ---------- Function checks empty lists & provided args ----------

def check_empty_lists(poster_index, list_1, list_2):
    if list_1 == [] or list_2 == []:
        print(f"{poster_index} error in dividing header")

            # ---------- Clean names from parasite symbols ----------

def clean_names(name):
    cleaned_name = re.sub(r"[†\*∗]", "", name)
    return cleaned_name

file = open("/Users/yevhenterziiason/Desktop/python_work/Friday_AB/sri_25_abstracts.txt", "r", encoding="UTF8")

text = []
posters = []

            # ---------- Separate index and clutter ----------

for line in file:
    if re.search(r"^\d{1,}[A-Z]?\s+Reproductive Sciences Vol\. 32, Supplement 1, March 2025 Scientific Abstracts", line.strip()):
        pass
    elif re.match(r"^((.*)-\d{3})$", line.strip()):
        posters.append(text)
        text = [line.strip()]
    else:
        if line.strip() != "":
            text.append(line.strip())

posters.append(text)

for poster in posters[1:]:
    abstract_flag = False
    header_raw = []
    poster_raw = []

            # ---------- Separate_header ----------

    for element in poster:
        if re.match(r"(Introduction:|Objective:)", element):
            abstract_flag = True
        if abstract_flag:
            poster_raw.append(element)
        else:
            header_raw.append(element)
    
    # check_empty_lists(element[0], header_raw, poster_raw)
    
    poster_index = header_raw[0]
    header_text = " ".join(header_raw[1:]).strip()
    poster_text = " ".join(poster_raw).strip()
    
    split_title = re.split(r'(\.|\?)\s', header_text, maxsplit=1)
    # if len(split_title) < 3:
    #     print(f"{poster_index}: can't split title properly")
    #     continue
    title = split_title[0] + "."
    authors_and_affil = split_title[2]
    
            # ---------- Matching single affiliation authors ----------

    if not re.search("1", authors_and_affil):
            authors_aff_separated = authors_and_affil.strip(".").rsplit(".", 1)

            single_aff_authors = authors_aff_separated[0]
            single_affiliation = authors_aff_separated[1].strip()
            
            authors = [a.strip() for a in re.split(r",", single_aff_authors) if a.strip()]
            for sa_author in authors:
                
                cleaned_names = clean_names(sa_author)
                
                ws.append([
                    cleaned_names,
                    single_affiliation,
                    role0,
                    email,
                    poster_index,
                    session_description,
                    title,
                    poster_text,
                    abstractUrl,
                    "",
            ])
            # ---------- Matching multiple affiliation authors ----------
    else:
        ma_authors_separated = re.split(r"(?<=\d)\s(?=\d)", authors_and_affil, maxsplit=1)
        split_authors = re.split(r"(?<=\d)\s", ma_authors_separated[0])
        split_affiliations = re.split(r"\s(?=\d)", ma_authors_separated[1])

        authors_with_numbers = []
        affiliations_with_numbers = []

        for author in split_authors:
            name = re.sub(r"[\d,†∗]", "", author).strip(" .")
            indicies = re.findall(r"\d+", author)
            authors_with_numbers.append((name, indicies))

        for affiliation in split_affiliations:
            clean_aff = re.sub(r"^\d+", "", affiliation).strip()
            index = re.findall(r"\d+", affiliation)
            affiliations_with_numbers.append((clean_aff, index))
        
            # ---------- Matching authors and affiliations ----------

        matched = []

        for author, author_inds in authors_with_numbers:
            affs = []
            for aff, ind_list in affiliations_with_numbers:
                if any(index in author_inds for index in ind_list):
                    affs.append(aff)
            matched.append({"name": author, "affiliation": " ___ ".join(affs)})

        for m in matched:
            
            cleaned_names = clean_names(m["name"])
            
            ws.append([
                cleaned_names,
                m["affiliation"],
                role0,
                email,
                poster_index,
                session_description,
                title,
                poster_text,
                abstractUrl,
                "",
            ])

wb.save("Friday_AB.xlsx")

        
        # print(poster_index, affiliations_with_numbers)

    #_____Split_authors_and_aff__________

        
        # raw_ma_authors_with_numbers = ma_authors_separated[0].strip()
        # if len(split_ma_authors) < 2:
        #     print(f"{poster_index}: multiple affiliation split failed")
        
        # for ma_author in ma_authors_separated:
        #     auth_name = re.sub(r"[\d,†∗]", "", ma_author).strip(" .")
        #     auth_number = re.findall(r"\d+", ma_author)
        #     ma_authors_with_numbers.append((auth_name, auth_number))
            
        #     raw_ma_affiliations_with_numbers = ma_authors_separated[1].strip()
        #     ma_affiliations_separated = re.split(r";\s*", raw_ma_affiliations_with_numbers.strip())
            
        # print(poster_index, ma_authors_with_numbers)
            
            



                
                
                # raw_ma_authors_with_numbers = ma_authors_separated[0].strip()
                # ma_affiliations_with_numbers = ma_authors_separated[1].strip()
            
            # ma_authors_raw = re.split(r"(?<=\d)\s(?=[A-Z])", ma_authors_with_numbers)
            # ma_affs_raw = re.split()
            # print(poster_index, ma_authors_with_numbers)
            
# __________________

#         # Build dict of affiliations
#         aff_dict = {}
#         for aff in re.split(r';\s(?=\d)', multiple_aff_affiliations):
#             match = re.match(r"(\d+)\s*(.*)", aff.strip())
#         if match:
#             num, institution = match.groups()
#             aff_dict[num] = institution.strip()

#         # Assign affiliations to authors
#         for i, author in enumerate(authors):
#             # Capture all numbers associated with the author
#             numbers = re.findall(r'\d+', author)

#             # Clean author name
#             clean_name = re.sub(r'[\*\†,\.;]?\d+', '', author).strip()

#             # Collect affiliations
#             aff_list = [aff_dict[num] for num in numbers if num in aff_dict]
#             aff_joined = " ___ ".join(aff_list) if aff_list else ""

#             ws.append([
#                 clean_name,
#                 aff_joined,
#                 role0,
#                 email,
#                 "",
#                 poster_index,
#                 title,
#                 poster_text,
#                 abstractUrl,
#                 "",
#             ])



# _____________unfihisned______________________________________
#     if re.findall(r"\d", authors_and_affil):
        
#         authors = re.split(r"(?<=\d)\s", split_authors[0])
#         aff_text = split_authors[1] 
#     else:
#         split_authors = re.split(r"\.\s", authors_and_affil, maxsplit=1)
#         authors = re.split(r",\s*", split_authors[0])
#         aff_text = split_authors[1]
#         print("\n", )

# _____________unfinished_______________________


# _______________

            # # def parse_authors_and_affil(authors_and_affil, index=None):
    # #         authors_with_numbers = []
    # #         affiliations_with_numbers = []
