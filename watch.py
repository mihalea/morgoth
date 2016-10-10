#!/usr/bin/env python3
import sys
import imaplib
import smtplib
import email
import email.header
import logging
import re
import time

import config as CFG

# Needed to extract the UID from the email
pattern_uid = re.compile('\d+ \(UID (?P<uid>\d+)\)')
logger = logging.getLogger("morgoth")


def setup_logger():
    """Setup the logging service with handlers for
    file logging and console logging
    """

    logger.setLevel(logging.DEBUG)
    # create file handler which logs even debug messages
    fh = logging.FileHandler('morgoth.log')
    fh.setLevel(logging.DEBUG)
    # create console handler with a higher log level
    ch = logging.StreamHandler()
    ch.setLevel(logging.DEBUG)
    # create formatter and add it to the handlers
    formatter = logging.Formatter('%(asctime)s - %(name)s - ' +
                                  '%(levelname)s - %(message)s')
    ch.setFormatter(formatter)
    fh.setFormatter(formatter)
    # add the handlers to logger
    logger.addHandler(ch)
    logger.addHandler(fh)


def connect():
    """ Connect to the SMTP server and return the connection """
    mail = imaplib.IMAP4_SSL(CFG.IMAP, CFG.IMAP_PORT)

    try:
        rv, data = mail.login(CFG.USER, CFG.PASS)
    except imaplib.IMAP4.error:
        logger.error("Login failed")
        sys.exit(1)

    rv, mailboxes = mail.list()
    if rv != "OK":
        logger.error("Failed to get mailboxes")
        sys.exit(1)

    rv, data = mail.select("INBOX")
    if rv != "OK":
        logger.error("Failed to select inbox")
        sys.exit(1)

    logger.info("Connected successfuly to the server")
    return mail


def forward(mail, num):
    """ Forward the given email to the configured email address """
    rv, data = mail.fetch(num, '(RFC822)')

    if rv != "OK" or data is None:
        logger.error("Failed to fetch email")
        return

    msg = email.message_from_bytes(data[0][1])
    msg.replace_header("To", CFG.TO)

    smtp = smtplib.SMTP(CFG.SMTP, CFG.SMTP_PORT)
    smtp.starttls()
    smtp.login(CFG.USER, CFG.PASS)
    smtp.sendmail(CFG.FROM, CFG.TO, msg.as_string())
    smtp.quit()

    logger.info("Forwarded mail: {}".format(num))


def archive(mail, num):
    """ Move the given email to the Archived folder """

    rv, data = mail.fetch(num, '(UID)')
    if rv != 'OK' or data is None:
        logger.error("Failed to fetch mail")
        return

    uid = parse_uid(data[0].decode("UTF-8"))
    rv, data = mail.uid('COPY', uid, 'Archived')
    if rv == 'OK':
        rv, data = mail.uid('STORE', uid, '+FLAGS', '(\Deleted)')
        mail.expunge()
        logger.info("Archived email: {}".format(num))


def parse_uid(data):
    ''' Extract to UID from the given string using regex '''
    match = pattern_uid.match(data)
    return match.group('uid')


def forward_and_archive(mail):
    """ Search from emails sent from the camera and then forward them
    to the configured email and archive them """

    rv, data = mail.search(None, 'FROM', '"sauron@mihalea.ro"')
    """ Number of emails archived so far. This is needed because for
    every email archived the index of the following emails decreases """
    archived = 0
    if rv == 'OK':
        for num in data[0].split():
            r_num = (str(int(num.decode("UTF-8")) - archived)).encode("UTF-8")
            logger.debug("Adjusting for deletions: {} -> {}"
                         .format(num, r_num))

            forward(mail, r_num)
            archive(mail, r_num)

            archived += 1

    logger.info("Processed {} emails".format(archived))


def main():
    """ Try to connect and process the email box """
    setup_logger()

    logger.info("Starting run loop")
    try:
        mail = connect()
        while mail is not None:
            forward_and_archive(mail)

            time.sleep(CFG.DELAY)
    except KeyboardInterrupt:
        logger.info("Closing by user request")
        sys.exit()
    finally:
        if mail is not None:
            logger.debug("Closing the email connection")
            mail.close()

if __name__ == '__main__':
    main()
