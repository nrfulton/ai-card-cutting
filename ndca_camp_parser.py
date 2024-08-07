import sys
import os
import jsonlines
from docx import Document
from card import Card, TAG_NAME
import argparse
from typing import List
import tqdm
import pickle
import glob
import hashlib

def cards_to_dict(cards):
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
  return json_dict


def parse_cards(filename: str, additional_info) -> List[Card]:
  document = Document(filename)
  cards = []
  current_card = []
  print("Parsing " + filename)
  for paragraph in document.paragraphs:
    if paragraph.style.name == TAG_NAME:
      try:
        cards.append(Card(current_card, additional_info=additional_info))
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
  parser.add_argument("directory", type=str, help="path to a directory of directories, each of which should contain docx files. e.g., files/Emory/Aff1.docx")
  parser.add_argument("-p", "--previous", type=str, help="previous parse results (for updating the file)")
  parser.add_argument("-o", "--output", type=str, help="path to output file (default: output.json)", default="TEMPORARY.jsonl")

  args = parser.parse_args()

  # Get list of files to parse
  files_to_parse = glob.glob(args.directory + "/*/*")

  files_to_skip = set()
  if args.previous:
    files_to_skip = set(
      map(
        lambda x: x['additional_info']['md5sum'],
        jsonlines.Reader(open(args.previous))
      )
    )
    print(f"skipping {len(files_to_skip)} files.")
  
  # Parse each file into a list of cards
  cards = []
  for source_directory, docx_filename in [(path[0], path[1]) for path in map(lambda y: y.replace(args.directory, "").split(os.sep), files_to_parse)]:
    docx_path = args.directory + os.sep + source_directory + os.sep + docx_filename
    try:
        hash = hashlib.md5(open(docx_path, 'rb').read()).hexdigest()
        if hash in files_to_skip:
          print(f"Skipping {docx_path}")
          continue
        parsed_cards = parse_cards(docx_path, 
                                   additional_info={
                                     "filename": docx_filename, 
                                     "md5sum": hash, 
                                     "camp_or_other_source": source_directory
                                  })
        for i, c in enumerate(parsed_cards):
          c["additional_info"]["order"] = i+1
        cards.extend(parsed_cards)
    except Exception as e:
        print(f"Error parsing {docx_path} from {source_directory}.")
        continue
    with open("cache.pickle", "wb") as fh:
      pickle.dump(cards_to_dict(cards), fh)
      print(f"saved {len(cards)} to the processing cache.")

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
  
  if args.previous:
    json_dict = list(jsonlines.Reader(open(args.previous)))
    json_dict.extend(cards_to_dict(cards))
    print(f"\t...but actually writing {len(json_dict)} files because we're combining with {args.previous}.")
  else:
    json_dict = cards_to_dict(cards)

  output_file = args.output
  with jsonlines.Writer(open(output_file, 'w')) as fh:
    fh.write_all(json_dict)

