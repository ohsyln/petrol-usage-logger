import gspread
import json
import logger
import requests
import time
from inbox import Inbox
from oauth2client.service_account import ServiceAccountCredentials

log = logger.Log(__file__)

############ GSPREAD ############
GAPI_CREDENTIALS = "account_key.json"
class SheetsAPI():
  def __init__(self):
    # Use Google Sheets API's credentials.json to access services
    scope = ['https://spreadsheets.google.com/feeds','https://www.googleapis.com/auth/drive']
    creds = ServiceAccountCredentials.from_json_keyfile_name(GAPI_CREDENTIALS, scope)
    client = gspread.authorize(creds)
    self.sheet = client.open('Finance').worksheet("PetrolSF")

  # Returns next available empty row
  def next_available_row(self):
    str_list = list(filter(None, self.sheet.col_values(1)))
    return len(str_list)+1

  # Helper function to create row for Gspread's append_row()
  def row_maker(self, date, mileage, refilled, costperlitre):
    row = [date, mileage, refilled, costperlitre]
    return row

  # Actual row update via Google Sheets API
  def update_row(self, date, mileage, refilled, costperlitre):
    rn = self.next_available_row()
    row = self.row_maker(date, mileage, refilled, costperlitre)
    self.sheet.append_row(row)
    self.sheet.update_acell("E{}".format(rn),"=C{0}*D{0}*0.84".format(rn))
    self.sheet.update_acell("G{}".format(rn),"=(B{0}-B{1})/C{0}".format(rn, rn-1))

class ParsingException(Exception):
  pass
class Parser():
  def __init__(self):
    pass
  # Explicit parsing of CaltexGO's receipt email body
  # Returns a 3-tuple: 
  #   (string) date
  #   (float) refilled
  #   (float) costperlitre
  def extract_info(self, body):
    main_body = self.get_text_between(body, "Volume:", "Amount")
    datetime_text = self.get_text_between(body, "Date & Time:", "Station:")
    try:
      [transaction_date, transaction_time] = datetime_text.split(",")
      [yyyy, mm, dd] = transaction_date.split("-")
      ddmmyy = dd+mm+yyyy[-2:]
    except:
      raise ParsingException("parsing error: can't parse datetime")

    delim = "litre @"
    if delim not in body:
      raise ParsingException("parsing error: can't find '{}' delim".format(delim))
    [refilled, costperlitre] = main_body.split(delim)
    refilled = refilled.strip()
    costperlitre = costperlitre.strip()
    try:
      refilled = float(refilled)
      costperlitre = float(costperlitre)
    except:
      raise ParsingException("parsing error: can't convert values to float")

    return ddmmyy, refilled, costperlitre
  # Helper function to get extract text between two specified keywords
  # In: 
  #   (string) body
  #   (string) first keyword
  #   (string) second keyword
  # Returns: 
  #   (string) text between the two keywords
  def get_text_between(self, s, start, end):
    s_i = s.find(start)
    e_i = s.find(end)
    if s_i == -1 or e_i == -1:
      raise ParsingException("parsing error: can't find '{}' or '{}'".format(start, end))
    return s[s_i + len(start): e_i].strip()

############ TELEGRAM BOT ############
API_KEY = "<YOUR TELEGRAM BOT API KEY HERE>"
API_PREFIX = 'https://api.telegram.org/bot'
API_SEND = 'sendMessage'
API_GET = 'getUpdates'
API_CHAT_ID = '<YOUR TELEGRAM BOT CHAT ID HERE>'
API_TEXT = '[{}] Please enter your mileage for petrol pump:'
class TelegramAPI():
  def __init__(self):
    pass
  # Sends prompter to user for mileage via Telegram bot
  # Returns (int) mileage
  def prompt_for_mileage(self):
    not_success = 1
    URL = API_PREFIX + API_KEY + API_SEND
    TEXT = API_TEXT.format(logger.now())
    PARAMS = {'chat_id': API_CHAT_ID, 'text': TEXT}
    while not_success:
      try:
        r = requests.get(url=URL, params=PARAMS)
      except requests.exceptions.RequestException as e:
        log.plog("RequestException: {} (retry in 60s")
        time.sleep(60)
      else:
        # Successfully sent prompt to user
        not_success = 0

    # Wait for reply
    not_success = 1
    URL = API_PREFIX + API_KEY + API_GET
    while not_success:
      try:
        r = requests.get(url=URL)
        mileage = self.parse_response(r)
      except requests.exceptions.RequestException as e:
        log.plog("RequestException: {} (retry in 60s")
        time.sleep(60)
      else:
        # If mileage exists, we proceed
        if mileage: 
          not_success = 0
        else:
          print("Waiting for user to enter mileage..")
          time.sleep(60)

    return mileage

  # Helper function to update latest mileage
  def update_mileage(self, m):
    f = open('mileage.txt', 'w')
    f.write(str(m))
    f.close()
  # Helper function to return (str) mileage for comparison
  def get_mileage(self):
    f = open('mileage.txt','r')
    mileage = f.read()
    f.close()
    return int(mileage)

  # Parse JSON response from user
  # Returns none if user has not sent their mileage via Telegram,
  # else returns user-input mileage
  def parse_response(self, r):
    j = json.loads(r.text)
    msg_list = j["result"]
    while len(msg_list) == 0:
      time.sleep(60)
    try:
      new_mileage = int(msg_list[-1]["message"]["text"])
    except:
      raise ParsingException("can't parse mileage from user: {}".format(msg_list[-1])) 
    if new_mileage > self.get_mileage():
      # User has provided an increased mileage -> re-fuelled
      self.update_mileage(new_mileage)
      return new_mileage
    return None

################## MAIN ####################33
def main():
  sheets_api = SheetsAPI()
  inbox = Inbox()
  parser = Parser()
  telegram = TelegramAPI()

  # Async callback when email is received
  @inbox.collate
  def handle(to, sender, body):
    str_body = body.decode("utf-8")
    # Caltex Singapore's receipt email body for CaltexGO
    if ("Thank You - Successful Payment (" not in str_body):
      # Ignore emails if conditions not met
      # e.g. using a very specific toaddr like ruioeqr1u138ry1@yourdomain.com
      print(str_body)
      return
    try:
      # Parses the fixed-format email to extract date, refill, cost per litre
      ddmmyy, refilled, costperlitre = parser.extract_info(str_body)
    except ParsingException as e:
      log.plog("{} (sender: {})".format(e, sender))
      return

    # Uses Telegram bot to prompt user to enter current mileage
    try:
      mileage = telegram.prompt_for_mileage()
    except ParsingException as e:
      print(e)
      return

    # Uses Gspread to access Google Sheets API to update cells
    sheets_api.update_row(ddmmyy, mileage, refilled, costperlitre)
    log.plog("update_row: {} {} {} {}".format(ddmmyy, mileage, refilled, costperlitre))

  inbox.serve(address='0.0.0.0', port=4467)

if __name__=="__main__":
  main()
