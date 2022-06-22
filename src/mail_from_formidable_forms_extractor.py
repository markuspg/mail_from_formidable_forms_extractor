#!/usr/bin/env python3

# mail_from_formidable_forms_extractor - tool to extract registrations from
# e-mail notifications sent by a "Formidable Forms" form
#
# Copyright (C) 2022 Markus Prasser
#
# This program is free software: you can redistribute it and/or modify it under
# the terms of the GNU General Public License as published by the Free Software
# Foundation, either version 3 of the License, or (at your option) any later
# version.

# This program is distributed in the hope that it will be useful, but WITHOUT
# ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS
# FOR A PARTICULAR PURPOSE.  See the GNU General Public License for more
# details.

# You should have received a copy of the GNU General Public License along with
# this program.  If not, see <https://www.gnu.org/licenses/>.

"""Tool to extract registrations from e-mail notifications sent by Formidable
Forms

"""

import csv
import email
import re
import ssl

from argparse import ArgumentParser
from email import policy
from getpass import getpass
from imaplib import IMAP4_SSL

# Global constants
EXTRACTION_REGEXP = re.compile(
    r"X-Mailer: WPMailSMTP/Mailer/smtp \d\.\d\.\d"
    r"\nMIME-Version: \d\.\d"
    r"\nContent-Type: text\/plain; charset=UTF-8"
    r"(\nContent-Transfer-Encoding: 8bit)?"
    r"\n"
    r"\n"
    r"\"eintrags_id\",\"key\",\"anzahl_personen\",\"zahlungsstatus\","
    r"\"name\",\"vorname\",\"mail\",\"telefon\",\"sonstiges\","
    r"\"club\",\"teilnehmer_2\",\"teilnehmer_3\",\"teilnehmer_4\","
    r"\"teilnehmer_5\",\"teilnehmer_6\",\"teilnehmer_7\","
    r"\n\"(?P<eintrags_id>\d+)\",\"(?P<key>[^\"]+)\","
    r"\"(?P<anzahl_personen>\d+)\",\"(?P<zahlungsstatus>[^\"]*)\","
    r"\"(?P<name>[^\"]+)\",\"(?P<vorname>[^\"]+)\",\"(?P<mail>[^\"]+)\","
    r"\"(?P<telefon>[^\"]*)\",\"(?P<sonstiges>[^\"]*)\","
    r"\"(?P<club>[^\"]*)\","
    r"\"(?P<teilnehmer_2>[^\"]*)\",\"(?P<teilnehmer_3>[^\"]*)\","
    r"\"(?P<teilnehmer_4>[^\"]*)\",\"(?P<teilnehmer_5>[^\"]*)\","
    r"\"(?P<teilnehmer_6>[^\"]*)\",\"(?P<teilnehmer_7>[^\"]*)\""
)
MAILBOX_FOR_DOWNLOAD = 'Registrations'

class RegistrationData: # pylint: disable=too-many-instance-attributes
    """Class representing the extracted data from one registration e-mail

    """
    def __init__(self,  data):
        """Initialize a new instance from a dictionary of the e-mail's data

        """
        self.eintrags_id = data['eintrags_id']
        self.key = data['key']
        self.anzahl_personen = data['anzahl_personen']
        self.zahlungsstatus = data['zahlungsstatus']
        self.name = data['name']
        self.vorname = data['vorname']
        self.mail = data['mail']
        self.telefon = data['telefon']
        self.sonstiges = data['sonstiges']
        self.club = data['club']
        self.teilnehmer_2 = data['teilnehmer_2']
        self.teilnehmer_3 = data['teilnehmer_3']
        self.teilnehmer_4 = data['teilnehmer_4']
        self.teilnehmer_5 = data['teilnehmer_5']
        self.teilnehmer_6 = data['teilnehmer_6']
        self.teilnehmer_7 = data['teilnehmer_7']

    @staticmethod
    def write_header_to_csv(csv_obj):
        """Write an applicable header to an object returned by csv.writer

        """
        csv_obj.writerow(
            [
                'eintrags_id', 'key', 'anzahl_personen', 'zahlungsstatus',
                'name', 'vorname', 'mail', 'telefon', 'sonstiges', 'club',
                'teilnehmer_2', 'teilnehmer_3', 'teilnehmer_4',
                'teilnehmer_5', 'teilnehmer_6', 'teilnehmer_7'
            ]
        )

    def write_to_csv(self, csv_obj):
        """Write the instance data to an object returned by csv.writer

        """
        csv_obj.writerow(
            [
                self.eintrags_id, self.key, self.anzahl_personen,
                self.zahlungsstatus, self.name, self.vorname, self.mail,
                self.telefon, self.sonstiges, self.club, self.teilnehmer_2,
                self.teilnehmer_3, self.teilnehmer_4, self.teilnehmer_5,
                self.teilnehmer_6, self.teilnehmer_7
            ]
        )

def check_if_message_to_skip(msg):
    """Check if an e-mail message is to be skipped

    """
    return (msg['X-Envelope-From'] != '<some_address@example.org>') \
            or (msg['X-Envelope-To'] != '<another_address@example.org>') \
            or (msg['From'] != '<some_address@example.org>') \
            or (msg['To'] != 'another_address@example.org, further_address@example.org')

def retrieve_registrations_from_server(cmdline_args):
    """Retrieve registrations from server

    Seemingly irrelevant e-mails are being filtered out.
    """
    # Create list to collect all registrations
    registrations = list()

    # Create a safe default configuration for the SSL connect
    ssl_cntxt = ssl.create_default_context()
    # Establish a context for the connection to the mail server, ensuring
    # automatic logout
    with IMAP4_SSL(
            host=cmdline_args.server_url, port=cmdline_args.port,
            ssl_context=ssl_cntxt, timeout=10) as imap_conn:
        # Identify the client using a password
        imap_conn.login(
            cmdline_args.imap_user,
            getpass("Please enter the password of \"" + cmdline_args.imap_user + "\": ")
        )

        # Choose the mailbox to operate on and forbid modifications
        imap_conn.select(MAILBOX_FOR_DOWNLOAD, readonly=True)

        # Query the server for all messages in the selected mailbox
        _tmp, data = imap_conn.search(None, 'ALL')

        # Process all messages in the mailbox
        for idx in data[0].split():
            # Download the message with index "idx" entirely
            _typ, data = imap_conn.fetch(idx, '(RFC822)')
            # Filter for the actual message data
            raw_email = data[0][1]
            # Convert the e-mail's bytes to text by assuming a UTF-8 encoding
            decoded = raw_email.decode('utf-8')
            # Create an email.message instance from the data
            msg = email.message_from_string(decoded, policy=policy.strict)

            if check_if_message_to_skip(msg):
                print(
                    "Skipping e-mail from \"{0}\" to \"{1}\"".format(
                        msg['X-Envelope-From'], msg['X-Envelope-To']
                    )
                )
                continue

            res = EXTRACTION_REGEXP.search(str(msg))
            registrations.append(RegistrationData(res.groupdict()))

        # Close the currently selected mailbox
        imap_conn.close()

    return registrations

def parse_commandline_arguments():
    """Pass the commandline arguments passed to this script

    """
    parser = ArgumentParser(
        description="Extract registrations from Formidable Forms e-mail notifications"
    )

    parser.add_argument(
        '-p', '--port', default=993,
        help='the port to connect to the IMAP server', type=int
    )
    parser.add_argument(
        '-v', '--version', action='version', version='%(prog)s 0.1.0'
    )
    parser.add_argument(
        'imap_user', help='the user for authentication to the IMAP server'
    )
    parser.add_argument(
        'server_url', help='the server to download the messages from'
    )

    return parser.parse_args()

def write_registrations_to_csv(registrations, export_file_name):
    """Export a list of RegistrationData objects to a CSV file

    """
    with open(export_file_name, 'wt', newline='') as csv_file:
        csv_writer = csv.writer(csv_file)
        RegistrationData.write_header_to_csv(csv_writer)
        for entry in registrations:
            entry.write_to_csv(csv_writer)

def main():
    """Main method and entry point of mail_from_formidable_forms_extractor

    """
    args = parse_commandline_arguments()

    registrations = retrieve_registrations_from_server(args)

    write_registrations_to_csv(registrations, 'exported_registrations.csv')

if __name__ == "__main__":
    main()

# vim: tabstop=8 expandtab shiftwidth=4 softtabstop=4
