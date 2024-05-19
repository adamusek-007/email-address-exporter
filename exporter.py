"""Whole program for exporting email addresses from mailbox via IMAP"""

import csv
import imaplib
import email
import re
import socket
import sys


class EmailAddressExporter:
    """This class is exporting email addresses from the mailbox"""

    # INPUT FIELDS
    imap_server_address: str = ""
    imap_email_username: str = ""
    imap_email_password: str = ""
    imap_port: int = 993
    output_csv_filename: str = ""

    # PLACEHOLDERS
    extracted_emails: set = set()
    user_mailbox_folders_names: set = set()
    user_mailbox_folders_count: int = 0
    user_mailbox_mails_count: int = 0
    user_mailbox_selected_folders: set = set()
    current_mail_number: int = 0
    selected_headers: set = set()

    # DEFAULT HEADERS TO LOOK INTO IT
    headers = ["To", "From", "Cc"]

    # MESSAGES
    server_connection_success_msg: str = (
        "ⓘ Ustanowiono połączenie z serwerem pocztowym {}"
    )
    server_connection_error_msg: str = (
        "❌ Nie udało sie ustanowić połączenia z serwerem {}"
    )

    user_logon_error_msg: str = "❌ Nie udało się zalogować na konto użytkonika {}"
    user_logon_success_msg: str = "ⓘ Pomyślnie zalogowano na konto użytkownika {}"

    user_mailbox_folders_count_label: str = (
        "ⓘ Znaleziono {} folderów w skrzynce pocztowej"
    )
    user_mailbox_folders_not_found_label: str = (
        "❌ Nie znaleziono folderów w skrzynce pocztowej"
    )
    user_mailbox_mails_count_label: str = (
        "ⓘ Znaleziono {} wiadomości w wybranych folderach"
    )
    user_mailbox_mails_not_found_label: str = (
        "❌ Nie znaleziono wiadomości w skrzynce pocztowej"
    )

    folder_mails_count_label: str = "ⓘ Znaleziono {} wiadomości w folderze: {}"
    select_options_label: str = (
        "Wybierz numer aby odznaczyć z przeszukiwania (0 aby zatwierdzić): "
    )
    invalid_option_label: str = "Niepoprawna opcja proszę wprowadź numer z zakresu."
    invalid_input_label: str = "Niepoprawne wejście proszę wprowadź numer."
    savingToFileInformation: str = (
        "ⓘ Zakończono przeszukiwanie, rozpoczynam zapis do pliku."
    )
    select_multiple_folders_information: str = (
        "Wybierz foldery wpisując ich numery (domyślnie wszystkie wybrane) (0 aby zatwierdzić):"
    )
    select_multiple_headers_information: str = (
        "Wybierz nagłówki z których chcesz pobrać adresy wpisując ich numery (domyślnie wszystkie wybrane) (0 aby zatwierdzić):"
    )
    end_of_file_save_information: str = "✅ Zakończono zapis do pliku."

    break_msg: str = "❗Przerwano eksport adresów e-mail.❗"

    folder_name_groups_reg: str = r'\s"\."\s(.*)$'

    charset_for_decoding_and_encoding: str = "utf-8"

    def __init__(self) -> None:
        try:
            self.mail_connection = self.__set_mail_connection()
            self.__print_connection_establish_success_information()
            self.__login_to_account()
            self.__set_folders_informations()
            self.__print_folders_count()
            self.__set_folders_to_look_into()
            self.__set_header_types_to_look_into()
            self.__set_all_mails_count()
            self.__print_mails_count()
            self.__loop_over_folders()
            self.__print_saving_to_file_information()
            self.__create_csv_file()
            self.mail_connection.logout()
            print(self.end_of_file_save_information)
        except KeyboardInterrupt:
            print(self.break_msg)

    def print_menu(self, options, selected) -> None:
        """Prints interactable menu"""
        number = 0
        for i, option in enumerate(options, start=1):
            number += 1
            checkbox = "✅ " if i in selected else "❌ "
            print(f"{number} {checkbox} {option}")

    def select_option(self, options, selected) -> None:
        """Handles user input for selection"""
        self.print_menu(options, selected)
        while True:
            try:
                choice = input(self.select_options_label)
                if choice == "0":
                    break
                choice = int(choice)
                if choice in range(1, len(options) + 1):
                    if choice in selected:
                        selected.remove(choice)
                    else:
                        selected.add(choice)
                    self.print_menu(options, selected)
                else:
                    print(self.invalid_option_label)
            except ValueError:
                print(self.invalid_input_label)

    def __set_header_types_to_look_into(self) -> None:
        """Sets header types to look for"""
        print(self.select_multiple_headers_information)
        selected_headers = set(range(1, len(self.headers) + 1))
        self.select_option(self.headers, selected_headers)
        for index in selected_headers:
            self.selected_headers.add(self.headers[index - 1])

    def __set_folders_to_look_into(self) -> None:
        """Sets folders to look into"""
        print(self.select_multiple_folders_information)
        selected_folders = set(range(1, len(self.user_mailbox_folders_names) + 1))
        self.select_option(self.user_mailbox_folders_names, selected_folders)
        for index in selected_folders:
            self.user_mailbox_selected_folders.add(
                self.user_mailbox_folders_names[index - 1]
            )

    def __print_saving_to_file_information(self) -> None:
        """Prints information about saving to file"""
        print(self.savingToFileInformation)

    def __print_connection_establish_success_information(self) -> None:
        """Prints information about successfull connection to the IMAP server."""
        print(self.server_connection_success_msg.format(self.imap_server_address))

    def __print_mails_count(self) -> None:
        """Prints mails count"""
        if self.user_mailbox_mails_count == 0:
            print(self.user_mailbox_mails_not_found_label)
            exit(6)
        else:
            print(
                self.user_mailbox_mails_count_label.format(
                    self.user_mailbox_mails_count
                )
            )
            print()

    def __print_folders_count(self) -> None:
        """Prints folders count"""
        if self.user_mailbox_folders_count == 0:
            print(self.user_mailbox_folders_not_found_label)
            exit(5)
        else:
            print(
                self.user_mailbox_folders_count_label.format(
                    self.user_mailbox_folders_count
                )
            )

    def __get_elements_count(self, iterable) -> int:
        """Get folder mails count"""
        byte_string = iterable[0]
        mails_count_str = byte_string.decode(self.charset_for_decoding_and_encoding)
        return int(mails_count_str)

    def __set_all_mails_count(self) -> None:
        """Sets all mails count"""
        for folder_name in self.user_mailbox_selected_folders:
            folder_status, folder_mails = self.mail_connection.select(folder_name)
            if folder_status == "OK":
                self.user_mailbox_mails_count += self.__get_elements_count(folder_mails)

    def __set_mail_connection(self) -> imaplib.IMAP4_SSL | imaplib.IMAP4:
        """Sets mail connection to the server"""
        if self.imap_port == 993:
            return self.__get_connection_via_ssl()
        else:
            return self.__get_connection()

    def __get_connection_via_ssl(self) -> imaplib.IMAP4_SSL:
        """Establishes conneciton to the server with SSL protection"""
        try:
            mail_connection = imaplib.IMAP4_SSL(
                self.imap_server_address, self.imap_port
            )
            return mail_connection
        except socket.gaierror:
            print(self.server_connection_error_msg.format(self.imap_server_address))
            sys.exit(2)

    def __get_connection(self) -> imaplib.IMAP4:
        """Establishes connection to the server without SSL protection"""
        try:
            mail_connection = imaplib.IMAP4(self.imap_server_address, self.imap_port)
            return mail_connection
        except socket.gaierror:
            print(self.server_connection_error_msg.format(self.imap_server_address))
            sys.exit(3)

    def __login_to_account(self) -> None:
        """Logins to user account"""
        try:
            self.mail_connection.login(
                self.imap_email_username, self.imap_email_password
            )
        except imaplib.IMAP4.error:
            print(self.user_logon_error_msg.format(self.imap_email_username))
            exit(4)
        print(self.user_logon_success_msg.format(self.imap_email_username))

    def __loop_over_folders_names(self, folders) -> set:
        """Loops over mailbox folders to get folders names"""
        folders_names = set()
        for folder in folders:
            folder_name_with_prefix = folder.decode(
                self.charset_for_decoding_and_encoding
            )
            folder_name_groups = re.search(
                self.folder_name_groups_reg, folder_name_with_prefix
            )
            if folder_name_groups:
                folder_name = folder_name_groups.group(1)
                folder = [2]
                folder[0] = folder_name
                folders_names.add(folder_name)
        folders_names = sorted(folders_names)
        return folders_names

    def __set_folders_informations(self) -> None:
        """Sets folders infomations"""
        status, user_mailbox_folders = self.mail_connection.list()
        if status == "OK":
            self.user_mailbox_folders_names = self.__loop_over_folders_names(
                user_mailbox_folders
            )
            self.user_mailbox_folders_count = int(len(self.user_mailbox_folders_names))

    def __decode_imap_mime_header(self, encoded_header) -> str:
        """Decodes IMAP mime header"""
        decoded_header = ""
        parts = email.header.decode_header(encoded_header)
        for part, encoding in parts:
            if encoding is None:
                if isinstance(part, bytes):
                    part = part.decode(self.charset_for_decoding_and_encoding)
                decoded_header += part
            else:
                if isinstance(part, bytes):
                    decoded_header += part.decode(encoding)
                else:
                    decoded_header += part
        return decoded_header

    def __update_console_line(self, content) -> None:
        """Updates console line"""
        print(content, end="\r")
        sys.stdout.flush()

    def __update_status_bar(self) -> None:
        """Updates status bar based on current mail number"""
        self.current_mail_number += 1
        percentage = self.current_mail_number / self.user_mailbox_mails_count * 100
        percentage_str = f"{percentage:.2f}%"
        self.__update_console_line(" ⏩ Postęp " + percentage_str)

    def __loop_over_mail_headers(self, mail) -> None:
        """Looping over mail headers"""
        self.__update_status_bar()
        mail_format = "(RFC822)"
        status, mail_data = self.mail_connection.fetch(mail, mail_format)
        if status == "OK":
            try:
                mail_msg = email.message_from_bytes(mail_data[0][1])
                for header in self.selected_headers:
                    mail_header = mail_msg[header]
                    if mail_header:
                        self.__clear_and_add_email(mail_header)
            except UnboundLocalError:
                pass
        else:
            print("Error fetching the email.")

    def __clear_and_add_email(self, to_header) -> None:
        """Function for clearing email header"""
        addresses = to_header.split(",")
        for address in addresses:
            header_elements = self.__decode_imap_mime_header(address)
            header_elements = header_elements.replace("<", "")
            header_elements = header_elements.replace(">", "")
            header_elements = header_elements.replace('"', "")
            self.extracted_emails.add(header_elements.strip())

    def __loop_over_folder_mails(self) -> None:
        """Looping over folder mails"""
        status, folder_mails = self.mail_connection.search(None, "ALL")
        if status == "OK":
            folder_mails = folder_mails[0].split()
            for mail in folder_mails:
                self.__loop_over_mail_headers(mail)

    def __loop_over_folder(self, folder_name, folder_mails) -> None:
        """Looping over folder content"""
        folder_mails_count = self.__get_elements_count(folder_mails)
        if folder_mails_count != 0:
            print(self.folder_mails_count_label.format(folder_mails_count, folder_name))
            self.__loop_over_folder_mails()

    def __loop_over_folders(self) -> None:
        """Looping over selected mailbox folders"""
        for folder_name in self.user_mailbox_selected_folders:
            folder_status, folder_mails = self.mail_connection.select(folder_name)
            if folder_status == "OK":
                self.__loop_over_folder(folder_name, folder_mails)

    def __create_csv_file(self) -> None:
        """Creates CSV file"""
        with open(
            self.output_csv_filename,
            mode="w",
            newline="",
            encoding=self.charset_for_decoding_and_encoding,
        ) as csvfile:
            csv_writer = csv.writer(csvfile)

            for from_address in self.extracted_emails:
                csv_writer.writerow([from_address])


EmailAddressExporter()
