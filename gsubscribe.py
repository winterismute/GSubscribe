#!/usr/bin/env python

# Automatic subscription script by Rudi Bonfiglioli
# This simple script checks the inbox of your gmail account for messages with a particular subject, reads sender address and the first two lines
#    then it opens a spreadsheet with a particular name, already created in your google docs account, and stores the data.
# Tne spreadsheet NEEDS TO HAVE three columns called: 'name', 'address', 'comment' - DO NOT USE CAPITAL LETTERS!!

# REQUIRES: imaplib, elementtree, gdata API (which is considered deprecated, but for small things still works quite ok)
import imaplib
import email
try:
  from xml.etree import ElementTree
except ImportError:
  from elementtree import ElementTree
import gdata.docs
import gdata.docs.service
import gdata.spreadsheet.service
import re
import os

#   NEEDED PARAMETERS
# The triggerin subject
mysub = ''
# The gmail account name
accname = ''
# The gmail account password
accpass = ''
# The google docs account name
gdocsname = ''
# The google docs account password
gdocspass = ''
# The name of the spreadsheet (tha must contains the columns: name, address, comment)
spname = ''
# For a sort of debug version with stdout prints, set this to 1
debugprint = 0

names_toadd = []
comments_toadd = []
addresses_toadd = []

def extract_body(payload):
   if isinstance(payload,str):
      return payload
   else:
      return ''.join([extract_body(part.get_payload()) for part in payload])

def StringToDictionary(row_data):
   result = {}
   for param in row_data.split():
      name, value = param.split('=')
      result[name] = value
   return result

if __name__ == "__main__":
   if(debugprint > 0):
      print 'Starting the script'
   conn = imaplib.IMAP4_SSL("imap.gmail.com", 993)
   conn.login(accname, accpass)
   conn.select()
   if(debugprint > 0):
      print 'Connected!'
   typ, data = conn.search(None, 'UNSEEN')
   try:
      for num in data[0].split():
         typ, msg_data = conn.fetch(num, '(RFC822)')
         for response_part in msg_data:
            if isinstance(response_part, tuple):
               msg = email.message_from_string(response_part[1])
               subject=msg['subject'] 
               if (subject == mysub):
                  addresses_toadd.append(msg['From'])
                  payload=msg.get_payload()
                  body=extract_body(payload)
                  lines = body.split('\n')
                  if (len(lines) > 0):
                     names_toadd.append(lines[0])
                  if (len(lines) > 1):
                     comments_toadd.append(lines[1])
                  else:
                     comments_toadd.append('NO COMMENTS')
         typ, response = conn.store(num, '+FLAGS', r'(\Seen)')
   finally:
      try:
         conn.close()
      except:
         pass
      conn.logout()

   # Now write on the spreadsheet

   if (len(names_toadd) > 0):
      if (debugprint > 0):
         print 'Email(s) read. Now writing the data...'
      spr_client = gdata.spreadsheet.service.SpreadsheetsService()
      spr_client.email = gdocsname
      spr_client.password = gdocspass
      spr_client.source = 'DGDARC-UberScript-1'
      spr_client.ProgrammaticLogin()
      
      q = gdata.spreadsheet.service.DocumentQuery()
      q['title'] = spname
      q['title-exact'] = 'true'
      feed = spr_client.GetSpreadsheetsFeed(query=q)
      sid = feed.entry[0].id.text.rsplit('/', 1)[1]
      feed = spr_client.GetWorksheetsFeed(sid)
      wid = feed.entry[0].id.text.rsplit('/', 1)[1]
      
      if (debugprint > 0):
         print 'Spreadsheet id:'
         print sid
         print 'Worksheet id:'
         print wid

      for i in xrange(len(names_toadd)):
         dict = {}
         dict['name'] = names_toadd[i];
         dict['address'] = addresses_toadd[i]
         dict['comment'] = comments_toadd[i]
         entry = spr_client.InsertRow(dict, sid)
         # This, at lest, should be really logged to a file...
         if (debugprint > 0):
            if isinstance(entry, gdata.spreadsheet.SpreadsheetsList):
               print "Insertion name" + names_toadd[i] + " succeeded!"
            else:
               print "Insertion name" + names_toadd[i] + "failed!!"
