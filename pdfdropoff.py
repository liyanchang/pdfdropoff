#!/usr/bin/python

"""
Additional Install Instructions:

For use with Linux and LibreOffice:
http://www.togaware.com/linux/survivor/Convert_MS_Word.html

For use with OS X:
http://code.google.com/p/wkhtmltopdf/

"""

import commands
import doctest
import logging
import time
import platform
import os

l = logging.getLogger()
l.setLevel(logging.DEBUG)
#l.setLevel(logging.INFO)

# Wrappers for OS
def darwin_pdf(source):
  return qlmanage(source)

def linux_pdf(source):
  return libreOffice_Pdf(source)


# Functions
def qlmanage(source):
  source = source.replace(' ', '\\ ')
  source = source.replace('(', '\(')
  source = source.replace(')', '\)')
  logging.info("Using QLManage Preview")
  out = qlmanage_Preview(source)
  if (not out):
    logging.info("Fallback: Using QLManage Thumbnail")
    out = qlmanage_Thumbnail(source)
  return out

def qlmanage_Preview(source):
  """
  >>> qlmanage_Preview('/home/dchang/Dropbox/pdf/test/test.doc')
  sh: qlmanage: not found
  False
  """

  if (not os.path.exists('temp/')):
    logging.debug("Creating Temp foler")
    os.mkdir('temp')

  out = commands.getoutput('qlmanage -p '+ source + ' -o temp')
  logging.debug("QLManage Preview Output: :" + out)

  if out.find('public.html') == -1:
    logging.debug("Didn't seem to generate an html doc. Fail.")
    return False

  filename = os.path.basename(source)
  temppath = 'temp/'+filename+'.qlpreview/Preview.html'
  logging.debug("Temp file: " + temppath)
  out = commands.getoutput("sed -i -e 's/<body class=\"s0\"><div class=\"PageStyle\">/<body><div>/' "+temppath)
  logging.debug("SED Output: " + out)
  out = commands.getoutput('./wkhtmltopdf-0.9.9-OS-X.i368 ' + temppath + ' ' + source + '.pdf')
  logging.debug("WebKit2HTML Output: " + out)

  return True

def qlmanage_Thumbnail(source):
  """
  >>> qlmanage_Thumbnail('/home/dchang/Dropbox/pdf/test/test.doc')
  sh: qlmanage: not found
  True
  """
  out = commands.getoutput('qlmanage -t -s 3000 ' + source + ' -o temp')
  logging.debug("QL Thumbnail output: " + out)
  
  #Path here not yet complete

  return True


def libreOffice_Pdf(source):
  """
  
  1. Takes a file (must be full source, no relative paths)
  2. Opens it with LibreOffice (requires LibreOffice 3+ with a specific macro installed)
  3. PDF prints the output.

  You'll have to manually check the output of these files to actually test

  >>> libreOffice_Pdf('/home/dchang/Dropbox/pdf/test/test.doc')
  True
  >>> libreOffice_Pdf('/home/dchang/Dropbox/pdf/test/test.docx')
  True
  >>> libreOffice_Pdf('/home/dchang/Dropbox/pdf/test/test.xlsx')
  True
  >>> libreOffice_Pdf('/home/dchang/Dropbox/pdf/test/no-exist.xlsx')
  False
  """

  logging.info("Using LibreOffice")
  if (not os.path.isfile(source)):
    logging.error("Couldn't find the file: " + source)
    return False
  cmd = 'soffice -invisible -norestore "macro:///Standard.Module1.ConvertWordToPDF(%s)"' % source
  logging.debug("Using Command: "+ cmd)
  out = commands.getoutput(cmd)
  return True


def usage():
  print "Usage: python pdfdropoff.py"
  return

def info():  
  info_text = [("Author", "Liyan David Chang"), 
               ("Email", "liyanchang@mit.edu"),
               ("Version", 0.1),
               ("Description", "This script watches for new files in a specified folder and changes them into a PDF")]
               
  for t in info_text:
    print "%-20s: %s" % t

  return 

if __name__ == "__main__":
  # Known Issues:
  # When a user overwrites an existing file, the script does nothing
  # A bit slow on the linux side
  # The Mac side is fairly inaccurate (Looks worse than the Quick Look UI)
  #

  # Config
  # Your computer username, ie /users/USERNAME/home etc.
  user = 'dchang'
  # The folder inside Dropbox.
  watch_dir = 'pdf/drop/'
  # How long to wait between polling
  sleep_time = 1

  os_name = platform.system()
  if os_name == 'Linux':
    watch_dir = os.path.join("/home/%s/Dropbox" % user, watch_dir)
    pdf_function = linux_pdf
  elif os_name == 'Darwin':
    watch_dir = os.path.join("/Users/%s/Dropbox" % user, watch_dir)
    pdf_function = darwin_pdf
  else:
    print "Unsupported OS"
    exit(0)

  logging.debug('OS: ' + os_name)
  logging.debug('Folder: ' + watch_dir)
  if (not os.path.isdir(watch_dir)):
    print "Watch Directory Not Found. Please check the configuration"
    print "Using: %s" % watch_dir
    exit(0)
    
  before = dict ([(f, None) for f in os.listdir(watch_dir)])
  
  while True:
    time.sleep (sleep_time)
    
    after = dict ([(f, None) for f in os.listdir(watch_dir)])
    added = [f for f in after if not f in before]
    
    logging.debug("New File(s): " + str(added))
    
    for source in added:
      ext = source.rsplit('.', 1)[-1].lower()

      logging.debug("New File: " + source)
      logging.debug("Extension: " + ext)

      if ext == 'pdf':
        logging.debug("It's a PDF. EOM")
        break
      else:
        logging.info("New File: " + source)
        fullsource = os.path.join(watch_dir, source)
        logging.debug("New File: " + fullsource)
        pdf_function(fullsource)        
        logging.info("DONE\n")

    before = after

