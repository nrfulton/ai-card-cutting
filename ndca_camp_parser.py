import sys
import os
import jsonlines
from docx import Document
from card import Card, TAG_NAME
import argparse
from typing import List
import tqdm

CITE_NAME = "13 pt Bold"

#def parse_cites(filename):
#  document = Document(filename)
#  cites = []
#  print("Parsing " + filename)
#  print("Found " + str(len(document.paragraphs)) + " paragraphs")
#
#  for paragraph in document.paragraphs:
#    cite = paragraph.text
#    for r in paragraph.runs:
#      if CITE_NAME in r.style.name or (r.style.font.bold or r.font.bold):
#        cite = cite.replace(r.text, "**" + r.text + "**")
#    cites.append(cite)
#  
#  return cites

def parse_cards(filename: str, source: str) -> List[Card]:
  document = Document(filename)
  cards = []
  current_card = []
  print("Parsing " + filename)
  for paragraph in document.paragraphs:
    if paragraph.style.name == TAG_NAME:
      try:
        cards.append(Card(current_card, additional_info={'camp_or_other_source': source, 'filename': filename}))
      except Exception as e:
        continue
      finally:
        current_card = [paragraph]
    else:
      current_card.append(paragraph)
  print(f"Parsed {len(cards)} cards.")
  return cards

if __name__ == "__main__":
  # Parse command line arguments
  parser = argparse.ArgumentParser()
  parser.add_argument("directory", type=str, help="path to a directory of docx files.")
  parser.add_argument("-o", "--output", type=str, help="path to output file (default: output.json)", default="output.jsonl")

  args = parser.parse_args()

  # Get list of files to parse
  files_to_parse = {"root": []}
  assert os.path.isdir(args.directory), "Expected directory."
  for file_or_dir in os.listdir(args.directory):
    if os.path.isfile(file_or_dir):
       files_to_parse["root"].append(args.directory + "/" + file_or_dir)
    elif os.path.isdir(args.directory + "/" + file_or_dir):
        files_to_parse[file_or_dir] = []
        for file in os.listdir(args.directory + "/" + file_or_dir):
            assert not os.path.isdir(args.directory + "/" + file_or_dir + "/" + file)
            if file.endswith(".docx"):
                files_to_parse[file_or_dir].append(args.directory + "/" + file_or_dir + "/" + file)
    else:
        print(f"File not found: {file_or_dir}")
        sys.exit(1)
  
  # Parse each file into a list of cards
  cards = []
  for source_directory, docx_path in tqdm.tqdm([(key, value) for key, values in files_to_parse.items() for value in values]):  
    try:
        parsed_cards = parse_cards(docx_path, source=source_directory)
        cards.extend(parsed_cards)
    except Exception as e:
        print(f"Error parsing {docx_path} from {source_directory}.")
        continue

  print("Found " + str(len(cards)) + " total cards")

  # Strip punctuation and empty strings from each card's highlighted text
  for card in cards:
    punctuation_list = [",", ".", "!", "?", ":", ";", "(", ")", "[", "]", "{", "}", "\"", "\'", "“", "”", "‘", "’"]

    # Strip punctuation from each word in list of highlighted words
    card.highlighted_text = [word.strip("".join(punctuation_list)) for word in card.highlighted_text]
    card.underlined_text = [word.strip("".join(punctuation_list)) for word in card.underlined_text]
    card.emphasized_text = [word.strip("".join(punctuation_list)) for word in card.emphasized_text]

    # Remove empty strings
    card.highlighted_text = list(filter(None, card.highlighted_text))
    card.underlined_text = list(filter(None, card.underlined_text))
    card.emphasized_text = list(filter(None, card.emphasized_text))

    # Assert that run_text length == highlight/underline/emphasis length
    try:
      assert len(card.run_text) == len(card.highlight_labels) 
      assert len(card.run_text) == len(card.underline_labels) 
      assert len(card.run_text) == len(card.emphasis_labels)
    except AssertionError:
      print("Error parsing " + card.tag + ": run_text length does not match highlight/underline/emphasis length")

    # Remove empty strings from run_text (and the associated labels)
    # Keep track of indexes to remove from labels
    indexes_to_remove = []
    for i, word in enumerate(card.run_text):
      if word == "":
        indexes_to_remove.append(i)
    for i in sorted(indexes_to_remove, reverse=True):
      del card.run_text[i]
      del card.highlight_labels[i]
      del card.underline_labels[i]
      del card.emphasis_labels[i]

  # Strip \u2018 and \u2019 and \u2014 from each card's card_text, tag, and underlined_text
  for card in cards:
    card.card_text = card.card_text.replace("\u2018", "'").replace("\u2019", "'").replace("\u2014", "-")
    card.tag = card.tag.replace("\u2018", "'").replace("\u2019", "'").replace("\u2014", "-")
    card.underlined_text = [word.replace("\u2018", "'").replace("\u2019", "'").replace("\u2014", "-") for word in card.underlined_text]
    card.highlighted_text = [word.replace("\u2018", "'").replace("\u2019", "'").replace("\u2014", "-") for word in card.highlighted_text]
    card.emphasized_text = [word.replace("\u2018", "'").replace("\u2019", "'").replace("\u2014", "-") for word in card.emphasized_text]
    card.run_text = [word.replace("\u2018", "'").replace("\u2019", "'").replace("\u2014", "-") for word in card.run_text]
    
  # Write cards to JSON file
  json_dict = [{
    "tag": card.tag, 
    "text": card.card_text, 
    "highlights": card.highlighted_text, 
    "underlines": card.underlined_text,
    "emphasis": card.emphasized_text,
    "cite": card.cite,
    "cite_emphasis": card.cite_emphasis,
    "run_text": card.run_text,
    "highlight_labels": card.highlight_labels,
    "underline_labels": card.underline_labels,
    "emphasis_labels": card.emphasis_labels,
    "additional_info": card.additional_info
  } for card in cards]

  output_file = args.output

  with jsonlines.Writer(open(output_file, 'w')) as fh:
    fh.write_all(json_dict)
