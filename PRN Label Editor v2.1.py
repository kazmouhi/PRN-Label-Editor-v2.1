import customtkinter as ctk
from tkinter import filedialog, messagebox
from PIL import Image, ImageDraw, ImageFont
import pandas as pd
from datetime import datetime
import subprocess
import socket
import win32print
import os
import threading
from typing import Optional
import csv
import win32api
import logging
from logging.handlers import RotatingFileHandler

class PRNEditor:
    def __init__(self):
        self.setup_logging()  # Setup logging first
        self.setup_window()
        self.setup_mappings()
        self.setup_gui()
        self.is_printing = False
        self.create_and_open_default_csv()  # Call the method to create and open the default CSV file

    def setup_logging(self):
        """Setup logging configuration"""
        # Create log directory if it doesn't exist
        log_dir = "./log"
        if not os.path.exists(log_dir):
            os.makedirs(log_dir)
        
        # Configure logging
        log_file = os.path.join(log_dir, "logs.txt")
        
        # Create logger
        self.logger = logging.getLogger("PRNEditor")
        self.logger.setLevel(logging.INFO)
        
        # Create formatter
        formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
        
        # Create file handler with rotation (max 5MB per file, keep 5 backup files)
        file_handler = RotatingFileHandler(
            log_file, 
            maxBytes=5*1024*1024, 
            backupCount=5,
            encoding='utf-8'
        )
        file_handler.setFormatter(formatter)
        
        # Add handler to logger
        self.logger.addHandler(file_handler)
        
        # Log startup
        self.logger.info("PRN Label Editor started")
        self.logger.info(f"Computer name: {socket.gethostname()}")

    def create_and_open_default_csv(self):
        # Define the default CSV file name
        default_csv_file = "default_labels.csv"
        default_prn_file = "Label.prn"
        
        # Check if CSV file already exists and is accessible
        csv_exists = os.path.exists(default_csv_file)
        
        if csv_exists:
            try:
                # Try to open the file to check if it's accessible
                with open(default_csv_file, 'r', newline='') as csvfile:
                    reader = csv.reader(csvfile)
                    # Check if it has the required columns
                    header = next(reader, None)
                    if header and all(col in header for col in ['AFMPN', 'Country', 'QTY']):
                        # File exists and has the right format, just use it
                        self.csv_path.set(default_csv_file)
                        self.show_status(f"Using existing CSV file: {default_csv_file}")
                        self.logger.info(f"Using existing CSV file: {default_csv_file}")
                        return
            except (IOError, PermissionError):
                # File is locked by another process, ask user what to do
                self.logger.warning(f"CSV file {default_csv_file} is locked by another process")
                result = messagebox.askyesno(
                    "File in Use", 
                    f"The file {default_csv_file} is currently open in another application.\n\n"
                    "Do you want to continue with the existing file without opening it?",
                    icon=messagebox.WARNING
                )
                if result:
                    self.csv_path.set(default_csv_file)
                    self.show_status(f"Using existing CSV file: {default_csv_file}")
                    self.logger.info(f"Using existing CSV file: {default_csv_file}")
                    return
                # If user says no, we'll create a new file below
        
        # Create a new CSV file with the required columns
        try:
            with open(default_csv_file, 'w', newline='') as csvfile:
                fieldnames = ['AFMPN', 'Country', 'QTY']
                writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
                writer.writeheader()
            
            # Try to open the CSV file in the default spreadsheet application
            try:
                os.startfile(default_csv_file)
                self.show_status(f"Default CSV file created and opened: {default_csv_file}")
                self.logger.info(f"Default CSV file created and opened: {default_csv_file}")
            except Exception as e:
                # If we can't open it (might be locked), just use it without opening
                self.show_status(f"Default CSV file created: {default_csv_file}")
                self.logger.info(f"Default CSV file created: {default_csv_file}")
            
            self.csv_path.set(default_csv_file)
                
        except Exception as e:
            error_msg = f"Error creating CSV file: {str(e)}"
            self.show_status(error_msg, is_error=True)
            self.logger.error(error_msg)
            
        # Set Default PRN file
        try:
            self.prn_path.set(default_prn_file)
            self.show_status(f"Default PRN file: {default_prn_file}")
            self.logger.info(f"Default PRN file set: {default_prn_file}")
        except Exception as e:
            error_msg = f"Error setting Default PRN file: {str(e)}"
            self.show_status(error_msg, is_error=True)
            self.logger.error(error_msg)

    def setup_window(self):
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("blue")  # We'll customize colors manually
        self.root = ctk.CTk()
        self.root.title("APTIV PRN Label Editor")
        self.root.geometry("800x800")
        self.root.minsize(600, 700)

    def setup_mappings(self):
        self.day_map = {str(i): c for i, c in enumerate('123456789ABCDEFGHJKLMNPQRSTVWX', 1)}
        self.month_map = {'01': 'a', '02': 'b', '03': 'c', '04': 'd', '05': 'e',
                          '06': 'f', '07': 'g', '08': 'h', '09': 'j', '10': 'k',
                          '11': 'm', '12': 'n'}
        self.year_map = {'2019': '9', '2020': 'A', '2021': 'B', '2022': 'C',
                         '2023': 'D', '2024': 'E', '2025': 'F'}
        self.CountrieCode_map = {
        'Andorra': 'AD', 'United Arab Emirates': 'AE', 'Afghanistan': 'AF', 'Antigua and Barbuda': 'AG',
        'Anguilla': 'AI', 'Albania': 'AL', 'Armenia': 'AM', 'Angola': 'AO', 'Antarctic Treaty area': 'AQ',
        'Argentina': 'AR', 'American Samoa': 'AS', 'Austria': 'AT', 'Australia': 'AU', 'Aruba': 'AW',
        'Åland': 'AX', 'Azerbaijan': 'AZ', 'Bosnia and Herzegovina': 'BA', 'Barbados': 'BB',
        'Bangladesh': 'BD', 'Belgium': 'BE', 'Burkina Faso': 'BF', 'Bulgaria': 'BG', 'Bahrain': 'BH',
        'Burundi': 'BI', 'Benin': 'BJ', 'Saint Barthélemy': 'BL', 'Bermuda': 'BM', 'Brunei': 'BN',
        'Bolivia': 'BO', 'Caribbean Netherlands': 'BQ', 'Brazil': 'BR', 'The Bahamas': 'BS',
        'Bhutan': 'BT', 'Bouvet Island': 'BV', 'Botswana': 'BW', 'Belarus': 'BY', 'Belize': 'BZ',
        'Canada': 'CA', 'Cocos (Keeling) Islands': 'CC', 'Democratic Republic of the Congo': 'CD',
        'Central African Republic': 'CF', 'Republic of the Congo': 'CG', 'Switzerland': 'CH',
        'Ivory Coast': 'CI', 'Cook Islands': 'CK', 'Chile': 'CL', 'Cameroon': 'CM',
        'People\'s Republic of China': 'CN', 'Colombia': 'CO', 'Costa Rica': 'CR', 'Cuba': 'CU',
        'Cape Verde': 'CV', 'Curaçao': 'CW', 'Christmas Island': 'CX', 'Cyprus': 'CY',
        'Czech Republic': 'CZ', 'Germany': 'DE', 'Djibouti': 'DJ', 'Denmark': 'DK', 'Dominica': 'DM',
        'Dominican Republic': 'DO', 'Algeria': 'DZ', 'Ecuador': 'EC', 'Estonia': 'EE', 'Egypt': 'EG',
        'Western Sahara': 'EH', 'Eritrea': 'ER', 'Spain': 'ES', 'Ethiopia': 'ET', 'Finland': 'FI',
        'Fiji': 'FJ', 'Falkland Islands': 'FK', 'Federated States of Micronesia': 'FM',
        'Faroe Islands': 'FO', 'France': 'FR', 'Gabon': 'GA', 'United Kingdom': 'GB', 'Great Britain': 'GB', 'Grenada': 'GD',
        'Georgia': 'GE', 'French Guiana': 'GF', 'Guernsey': 'GG', 'Ghana': 'GH', 'Gibraltar': 'GI',
        'Greenland': 'GL', 'The Gambia': 'GM', 'Guinea': 'GN', 'Guadeloupe': 'GP',
        'Equatorial Guinea': 'GQ', 'Greece': 'GR', 'South Georgia and the South Sandwich Islands': 'GS',
        'Guatemala': 'GT', 'Guam': 'GU', 'Guinea-Bissau': 'GW', 'Guyana': 'GY', 'Hong Kong': 'HK',
        'Heard Island and McDonald Islands': 'HM', 'Honduras': 'HN', 'Croatia': 'HR', 'Haiti': 'HT',
        'Hungary': 'HU', 'Indonesia': 'ID', 'Ireland': 'IE', 'Israel': 'IL', 'Isle of Man': 'IM',
        'India': 'IN', 'British Indian Ocean Territory': 'IO', 'Iraq': 'IQ', 'Iran': 'IR',
        'Iceland': 'IS', 'Italy': 'IT', 'Jersey': 'JE', 'Jamaica': 'JM', 'Jordan': 'JO', 'Japan': 'JP',
        'Kenya': 'KE', 'Kyrgyzstan': 'KG', 'Cambodia': 'KH', 'Kiribati': 'KI', 'Comoros': 'KM',
        'Saint Kitts and Nevis': 'KN', 'North Korea': 'KP', 'South Korea': 'KR', 'Kuwait': 'KW',
        'Cayman Islands': 'KY', 'Kazakhstan': 'KZ', 'Laos': 'LA', 'Lebanon': 'LB',
        'Saint Lucia': 'LC', 'Liechtenstein': 'LI', 'Sri Lanka': 'LK', 'Liberia': 'LR', 'Lesotho': 'LS',
        'Lithuania': 'LT', 'Luxembourg': 'LU', 'Latvia': 'LV', 'Libya': 'LY', 'Morocco': 'MA',
        'Monaco': 'MC', 'Moldova': 'MD', 'Montenegro': 'ME', 'Saint-Martin': 'MF',
        'Madagascar': 'MG', 'Marshall Islands': 'MH', 'North Macedonia': 'MK', 'Mali': 'ML',
        'Myanmar': 'MM', 'Mongolia': 'MN', 'Macau': 'MO', 'Northern Mariana Islands': 'MP',
        'Martinique': 'MQ', 'Mauritania': 'MR', 'Montserrat': 'MS', 'Malta': 'MT',
        'Mauritius': 'MU', 'Maldives': 'MV', 'Malawi': 'MW', 'Mexico': 'MX', 'Malaysia': 'MY',
        'Mozambique': 'MZ', 'Namibia': 'NA', 'New Caledonia': 'NC', 'Niger': 'NE',
        'Norfolk Island': 'NF', 'Nigeria': 'NG', 'Nicaragua': 'NI', 'Kingdom of the Netherlands': 'NL',
        'Norway': 'NO', 'Nepal': 'NP', 'Nauru': 'NR', 'Niue': 'NU', 'New Zealand': 'NZ',
        'Oman': 'OM', 'Panama': 'PA', 'Peru': 'PE', 'French Polynesia': 'PF', 'Papua New Guinea': 'PG',
        'Philippines': 'PH', 'Pakistan': 'PK', 'Poland': 'PL', 'Saint Pierre and Miquelon': 'PM',
        'Pitcairn Islands': 'PN', 'Puerto Rico': 'PR', 'Palestine': 'PS', 'Portugal': 'PT',
        'Palau': 'PW', 'Paraguay': 'PY', 'Qatar': 'QA', 'Réunion': 'RE', 'Romania': 'RO',
        'Serbia': 'RS', 'Russia': 'RU', 'Rwanda': 'RW', 'Saudi Arabia': 'SA',
        'Solomon Islands': 'SB', 'Seychelles': 'SC', 'Sudan': 'SD', 'Sweden': 'SE',
        'Singapore': 'SG', 'Saint Helena, Ascension and Tristan da Cunha': 'SH',
        'Slovenia': 'SI', 'Svalbard and Jan Mayen': 'SJ', 'Slovakia': 'SK', 'Sierra Leone': 'SL',
        'San Marino': 'SM', 'Senegal': 'SN', 'Somalia': 'SO', 'Suriname': 'SR', 'South Sudan': 'SS',
        'São Tomé and Príncipe': 'ST', 'El Salvador': 'SV', 'Sint Maarten': 'SX', 'Syria': 'SY',
        'Eswatini': 'SZ', 'Turks and Caicos Islands': 'TC', 'Chad': 'TD',
        'French Southern and Antarctic Lands': 'TF', 'Togo': 'TG', 'Thailand': 'TH', 'Tajikistan': 'TJ',
        'Tokelau': 'TK', 'Timor-Leste': 'TL', 'Turkmenistan': 'TM', 'Tunisia': 'TN', 'Tonga': 'TO',
        'Turkey': 'TR', 'Trinidad and Tobago': 'TT', 'Tuvalu': 'TV', 'Taiwan': 'TW', 'Tanzania': 'TZ',
        'Ukraine': 'UA', 'Uganda': 'UG', 'United States Minor Outlying Islands': 'UM',
        'United States': 'US', 'Uruguay': 'UY', 'Uzbekistan': 'UZ', 'Vatican City': 'VA',
        'Saint Vincent and the Grenadines': 'VC', 'Venezuela': 'VE', 'British Virgin Islands': 'VG',
        'United States Virgin Islands': 'VI', 'Vietnam': 'VN', 'Vanuatu': 'VU',
        'Wallis and Futuna': 'WF', 'Samoa': 'WS', 'Yemen': 'YE', 'Mayotte': 'YT', 'South Africa': 'ZA',
        'Zambia': 'ZM', 'Zimbabwe': 'ZW'
        }

    def create_aptiv_logo(self, size=(500, 90)):
        """
        Create the official Aptiv logo with proper proportions
        """
        W, H = size
        img = Image.new('RGBA', size, (255, 255, 255, 255))  # White background
        draw = ImageDraw.Draw(img)

        # Calculate scaling based on width
        letter_height = H * 0.7  # Letters take up 70% of height
        letter_width = letter_height * 0.8  # Width-to-height ratio for letters
        spacing = letter_width * 0.3  # Space between letters
        
        # Starting position (centered horizontally)
        total_width = 5 * letter_width + 4 * spacing + letter_width * 0.5  # 5 letters + 4 spaces + dot space
        start_x = (W - total_width) / 2
        start_y = (H - letter_height) / 2
        
        # Draw the orange dots
        dot_radius = letter_height * 0.1
        dot_y = start_y + letter_height * 0.5
        
        # Left dot
        left_dot_x = start_x - dot_radius * 3
        draw.ellipse([
            (left_dot_x - dot_radius, dot_y - dot_radius),
            (left_dot_x + dot_radius, dot_y + dot_radius)
        ], fill=(255, 107, 0))
        
        # Right dot
        right_dot_x = start_x + total_width - letter_width * 0.5 + dot_radius * 3
        draw.ellipse([
            (right_dot_x - dot_radius, dot_y - dot_radius),
            (right_dot_x + dot_radius, dot_y + dot_radius)
        ], fill=(255, 107, 0))

        # Letter positions
        positions = []
        for i in range(5):
            x = start_x + i * (letter_width + spacing)
            positions.append(x)

        # Draw letters A P T I V
        stroke_width = max(2, int(letter_height / 20))
        
        # Letter A
        x = positions[0]
        # Left diagonal
        draw.line([(x + letter_width * 0.1, start_y + letter_height), 
                  (x + letter_width * 0.5, start_y)], fill=(0, 0, 0), width=stroke_width)
        # Right diagonal
        draw.line([(x + letter_width * 0.5, start_y), 
                  (x + letter_width * 0.9, start_y + letter_height)], fill=(0, 0, 0), width=stroke_width)
        # Horizontal bar
        draw.line([(x + letter_width * 0.25, start_y + letter_height * 0.6), 
                  (x + letter_width * 0.75, start_y + letter_height * 0.6)], fill=(0, 0, 0), width=stroke_width)

        # Letter P
        x = positions[1]
        # Vertical line
        draw.line([(x + letter_width * 0.1, start_y), 
                  (x + letter_width * 0.1, start_y + letter_height)], fill=(0, 0, 0), width=stroke_width)
        # Top horizontal
        draw.line([(x + letter_width * 0.1, start_y), 
                  (x + letter_width * 0.8, start_y)], fill=(0, 0, 0), width=stroke_width)
        # Middle horizontal
        draw.line([(x + letter_width * 0.1, start_y + letter_height * 0.5), 
                  (x + letter_width * 0.7, start_y + letter_height * 0.5)], fill=(0, 0, 0), width=stroke_width)
        # Right vertical (top half)
        draw.line([(x + letter_width * 0.8, start_y), 
                  (x + letter_width * 0.8, start_y + letter_height * 0.5)], fill=(0, 0, 0), width=stroke_width)

        # Letter T
        x = positions[2]
        # Top horizontal
        draw.line([(x, start_y), 
                  (x + letter_width, start_y)], fill=(0, 0, 0), width=stroke_width)
        # Vertical center
        draw.line([(x + letter_width * 0.5, start_y), 
                  (x + letter_width * 0.5, start_y + letter_height)], fill=(0, 0, 0), width=stroke_width)

        # Letter I
        x = positions[3]
        # Top horizontal
        draw.line([(x + letter_width * 0.2, start_y), 
                  (x + letter_width * 0.8, start_y)], fill=(0, 0, 0), width=stroke_width)
        # Vertical center
        draw.line([(x + letter_width * 0.5, start_y), 
                  (x + letter_width * 0.5, start_y + letter_height)], fill=(0, 0, 0), width=stroke_width)
        # Bottom horizontal
        draw.line([(x + letter_width * 0.2, start_y + letter_height), 
                  (x + letter_width * 0.8, start_y + letter_height)], fill=(0, 0, 0), width=stroke_width)

        # Letter V
        x = positions[4]
        # Left diagonal
        draw.line([(x + letter_width * 0.1, start_y), 
                  (x + letter_width * 0.5, start_y + letter_height)], fill=(0, 0, 0), width=stroke_width)
        # Right diagonal
        draw.line([(x + letter_width * 0.9, start_y), 
                  (x + letter_width * 0.5, start_y + letter_height)], fill=(0, 0, 0), width=stroke_width)

        return img

    def create_printer_icon(self, fill="white", size=(128, 128)):
        image = Image.new('RGBA', size, (0, 0, 0, 0))
        draw = ImageDraw.Draw(image)

        width, height = size

        # Enhanced printer design
        body_left = width * 0.1
        body_right = width * 0.9
        body_top = height * 0.2
        body_bottom = height * 0.8

        # Main printer body with rounded corners
        draw.rounded_rectangle(
            [(body_left, body_top), (body_right, body_bottom)],
            radius=width * 0.05,
            outline=fill,
            width=max(2, int(width / 32))
        )

        # Paper tray
        tray_height = height * 0.1
        draw.line(
            [(body_left, body_bottom - tray_height),
             (body_left - width * 0.05, body_bottom)],
            fill=fill,
            width=max(2, int(width / 32))
        )

        # Control panel
        panel_width = (body_right - body_left) * 0.4
        panel_height = (body_bottom - body_top) * 0.15
        panel_left = body_left + (body_right - body_left) * 0.1
        panel_top = body_top + (body_bottom - body_top) * 0.1

        draw.rectangle(
            [(panel_left, panel_top),
             (panel_left + panel_width, panel_top + panel_height)],
            outline=fill,
            width=max(2, int(width / 32))
        )

        # Buttons
        button_size = panel_height * 0.4
        for i in range(3):
            x = panel_left + button_size + (i * button_size * 1.5)
            y = panel_top + panel_height / 2
            draw.ellipse(
                [(x - button_size / 2, y - button_size / 2),
                 (x + button_size / 2, y + button_size / 2)],
                outline=fill,
                width=max(1, int(width / 64))
            )

        return image

    def setup_gui(self):
        self.prn_path = ctk.StringVar()
        self.csv_path = ctk.StringVar()
        self.computer_name = ctk.StringVar(value=socket.gethostname())
        self.status_message = ctk.StringVar(value="Ready to start")
        self.status_color = ctk.StringVar(value="#FF6B00")  # APTIV orange

        # APTIV Logo Header - Full width at the very top (no padding)
        logo_header = ctk.CTkFrame(self.root, fg_color="#ffffff", corner_radius=0, height=100)
        logo_header.pack(fill="x", side="top")
        logo_header.pack_propagate(False)
        
        # Center the logo in the header
        logo_container = ctk.CTkFrame(logo_header, fg_color="transparent")
        logo_container.pack(expand=True, fill="both")
        
        logo_label = ctk.CTkLabel(
            logo_container,
            text="",
            image=ctk.CTkImage(
                light_image=self.create_aptiv_logo(size=(500, 80)),
                dark_image=self.create_aptiv_logo(size=(500, 80)),
                size=(500, 80)
            )
        )
        logo_label.pack(expand=True)

        # Main container (below the logo)
        container = ctk.CTkFrame(self.root, fg_color="transparent")
        container.pack(fill="both", expand=True, padx=30, pady=20)

        # Header with title and computer info
        header = ctk.CTkFrame(container, fg_color="#1a1a1a", corner_radius=10)
        header.pack(fill="x", pady=(0, 20))
        
        # Header content frame
        header_content = ctk.CTkFrame(header, fg_color="transparent")
        header_content.pack(fill="x", padx=20, pady=10)
        
        # Title
        ctk.CTkLabel(
            header_content,
            text="PRN Label Editor",
            font=("Helvetica", 24, "bold"),
            text_color="#FF6B00"  # APTIV orange
        ).pack()

        # Status message (replacing popup messages)
        self.status_frame = ctk.CTkFrame(
            container,
            fg_color="#1a1a1a",
            corner_radius=10,
            height=60
        )
        self.status_frame.pack(fill="x", pady=(0, 20))
        self.status_frame.pack_propagate(False)

        self.status_label = ctk.CTkLabel(
            self.status_frame,
            textvariable=self.status_message,
            font=("Helvetica", 12),
            text_color="#ffffff"
        )
        self.status_label.pack(expand=True)

        # File selection section
        file_frame = ctk.CTkFrame(container, fg_color="#1a1a1a", corner_radius=10)
        file_frame.pack(fill="x", pady=(0, 20))

        # PRN selection
        prn_frame = ctk.CTkFrame(file_frame, fg_color="transparent")
        prn_frame.pack(fill="x", padx=20, pady=10)

        ctk.CTkButton(
            prn_frame,
            text="Select PRN",
            command=self.select_prn,
            height=36,
            width=120,
            fg_color="#FF6B00",  # APTIV orange
            hover_color="#FF8C3A",  # Lighter orange
            corner_radius=8
        ).pack(side="left", padx=(0, 10))

        ctk.CTkEntry(
            prn_frame,
            textvariable=self.prn_path,
            height=36,
            fg_color="#0d0d0d",
            border_color="#333333",
            corner_radius=8,
            placeholder_text="Select PRN template file..."
        ).pack(side="left", fill="x", expand=True)

        # CSV selection
        csv_frame = ctk.CTkFrame(file_frame, fg_color="transparent")
        csv_frame.pack(fill="x", padx=20, pady=10)

        ctk.CTkButton(
            csv_frame,
            text="Select CSV",
            command=self.select_csv,
            height=36,
            width=120,
            fg_color="#FF6B00",  # APTIV orange
            hover_color="#FF8C3A",  # Lighter orange
            corner_radius=8
        ).pack(side="left", padx=(0, 10))

        ctk.CTkEntry(
            csv_frame,
            textvariable=self.csv_path,
            height=36,
            fg_color="#0d0d0d",
            border_color="#333333",
            corner_radius=8,
            placeholder_text="Select CSV data file..."
        ).pack(side="left", fill="x", expand=True)

        # Settings section
        settings_frame = ctk.CTkFrame(container, fg_color="#1a1a1a", corner_radius=10)
        settings_frame.pack(fill="x", pady=(0, 20))

        # Computer name
        computer_frame = ctk.CTkFrame(settings_frame, fg_color="transparent")
        computer_frame.pack(fill="x", padx=20, pady=10)

        ctk.CTkLabel(
            computer_frame,
            text="Computer:",
            font=("Helvetica", 12, "bold"),
            text_color="#FFFFFF"
        ).pack(side="left", padx=(0, 10))

        ctk.CTkEntry(
            computer_frame,
            textvariable=self.computer_name,
            state='readonly',
            height=36,
            fg_color="#0d0d0d",
            border_color="#333333",
            corner_radius=8,
            text_color="#FFFFFF"
        ).pack(side="left", fill="x", expand=True)

        # Printer selection
        printer_frame = ctk.CTkFrame(settings_frame, fg_color="transparent")
        printer_frame.pack(fill="x", padx=20, pady=10)

        ctk.CTkLabel(
            printer_frame,
            text="Printer:",
            font=("Helvetica", 12, "bold"),
            text_color="#FFFFFF"
        ).pack(side="left", padx=(0, 10))

        self.printer_combo = ctk.CTkOptionMenu(
            printer_frame,
            values=self.get_printers(),
            height=36,
            fg_color="#FF6B00",  # APTIV orange
            button_color="#FF6B00",  # APTIV orange
            button_hover_color="#FF8C3A",  # Lighter orange
            corner_radius=8,
            text_color="#FFFFFF"
        )
        self.printer_combo.pack(side="left", fill="x", expand=True)

        # Control buttons frame
        control_frame = ctk.CTkFrame(container, fg_color="transparent")
        control_frame.pack(fill="x", pady=(0, 20))

        # Print button
        self.print_button = ctk.CTkButton(
            control_frame,
            text="Print Labels",
            command=self.start_printing,
            height=50,
            font=("Helvetica", 18, "bold"),
            compound="left",
            fg_color="#FF6B00",  # APTIV orange
            hover_color="#FF8C3A",  # Lighter orange
            corner_radius=10,
            image=ctk.CTkImage(
                light_image=self.create_printer_icon(fill="white"),
                dark_image=self.create_printer_icon(fill="white"),
                size=(32, 32)
            )
        )
        self.print_button.pack(side="left", expand=True, padx=5)

        # Cancel button (hidden by default)
        self.cancel_button = ctk.CTkButton(
            control_frame,
            text="Cancel",
            command=self.cancel_printing,
            height=50,
            font=("Helvetica", 18, "bold"),
            fg_color="#333333",  # Dark gray
            hover_color="#555555",  # Lighter gray
            corner_radius=10,
        )
        self.cancel_button.pack(side="left", expand=True, padx=5)
        self.cancel_button.pack_forget()

        # Progress section
        progress_frame = ctk.CTkFrame(container, fg_color="#1a1a1a", corner_radius=10)
        progress_frame.pack(fill="both", expand=True)

        self.current_label = ctk.CTkLabel(
            progress_frame,
            text="Current Label: -",
            font=("Helvetica", 12),
            text_color="#FFFFFF"
        )
        self.current_label.pack(pady=10)

        self.copies_label = ctk.CTkLabel(
            progress_frame,
            text="Copies: -",
            font=("Helvetica", 12),
            text_color="#FFFFFF"
        )
        self.copies_label.pack(pady=5)

        self.total_label = ctk.CTkLabel(
            progress_frame,
            text="Total Progress: -",
            font=("Helvetica", 12),
            text_color="#FFFFFF"
        )
        self.total_label.pack(pady=5)

        self.copies_progress = ctk.CTkProgressBar(
            progress_frame,
            height=15,
            corner_radius=5,
            fg_color="#0d0d0d",
            progress_color="#FF6B00"  # APTIV orange
        )
        self.copies_progress.pack(pady=10, padx=30, fill="x")
        self.copies_progress.set(0)

        self.total_progress = ctk.CTkProgressBar(
            progress_frame,
            height=15,
            corner_radius=5,
            fg_color="#0d0d0d",
            progress_color="#FF6B00"  # APTIV orange
        )
        self.total_progress.pack(pady=10, padx=30, fill="x")
        self.total_progress.set(0)

    def show_status(self, message: str, is_error: bool = False):
        self.status_message.set(message)
        color = "#FF3333" if is_error else "#FF6B00"  # Red for errors, orange for normal
        self.status_label.configure(text_color=color)
        
        # Log the status message
        if is_error:
            self.logger.error(message)
        else:
            self.logger.info(message)
            
        self.root.update()

    def get_printers(self) -> list:
        try:
            printers = []
            for printer in win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS):
                printers.append(printer[2])
            self.logger.info(f"Found {len(printers)} printers: {printers}")
            return printers if printers else ['No printers found']
        except Exception as e:
            error_msg = f"Error getting printers: {str(e)}"
            self.logger.error(error_msg)
            return ['No printers found']

    def select_prn(self):
        filename = filedialog.askopenfilename(filetypes=[("PRN files", "*.prn")])
        if filename:
            self.prn_path.set(filename)
            status_msg = f"Selected PRN file: {os.path.basename(filename)}"
            self.show_status(status_msg)
            self.logger.info(status_msg)

    def select_csv(self):
        filename = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])
        if filename:
            self.csv_path.set(filename)
            status_msg = f"Selected CSV file: {os.path.basename(filename)}"
            self.show_status(status_msg)
            self.logger.info(status_msg)

    def get_date_code(self) -> str:
        now = datetime.now()
        day_code = self.day_map[str(now.day)]
        month_code = self.month_map[f"{now.month:02d}"]
        year_code = self.year_map[str(now.year)]
        return f"{day_code}{month_code}{year_code}"

    def get_position(self, qty: int) -> str:
        if qty < 10:
            return "3101"
        elif qty < 100:
            return "2801"
        else:
            return "2501"

    def toggle_controls(self, printing: bool):
        if printing:
            self.print_button.pack_forget()
            self.cancel_button.pack(side="left", expand=True, padx=5)
        else:
            self.cancel_button.pack_forget()
            self.print_button.pack(side="left", expand=True, padx=5)

    def start_printing(self):
        if not self.validate_inputs():
            return

        self.is_printing = True
        self.toggle_controls(True)
        self.show_status("Starting print job...")
        self.logger.info("Print job started")

        # Start printing in a separate thread
        threading.Thread(target=self.process_and_print, daemon=True).start()

    def validate_inputs(self) -> bool:
        if not self.prn_path.get():
            error_msg = "Please select a PRN template file"
            self.show_status(error_msg, True)
            self.logger.warning(error_msg)
            return False

        if not self.csv_path.get():
            error_msg = "Please select a CSV data file"
            self.show_status(error_msg, True)
            self.logger.warning(error_msg)
            return False

        if not self.printer_combo.get() or self.printer_combo.get() == 'No printers found':
            error_msg = "Please select a valid printer"
            self.show_status(error_msg, True)
            self.logger.warning(error_msg)
            return False

        return True

    def cancel_printing(self):
        self.is_printing = False
        self.show_status("Print job cancelled")
        self.logger.warning("Print job cancelled by user")
        self.toggle_controls(False)

    def update_progress(self, current_label: str, copy_num: int, total_copies: int,
                        design_num: int, total_designs: int, total_labels_printed: int,
                        total_to_print: int):
        self.current_label.configure(
            text=f"Current Label: Design {design_num} of {total_designs} ({current_label})"
        )
        self.copies_label.configure(
            text=f"Copies: {copy_num} of {total_copies}"
        )
        self.total_label.configure(
            text=f"Total Progress: {total_labels_printed} of {total_to_print}"
        )

        self.copies_progress.set(copy_num / total_copies if total_copies > 0 else 0)
        self.total_progress.set(total_labels_printed / total_to_print if total_to_print > 0 else 0)
        self.root.update()

    def print_to_printer(self, prn_content: str, printer_name: str):
        """Print PRN content directly to the selected printer"""
        try:
            # Open the printer
            hprinter = win32print.OpenPrinter(printer_name)
            
            try:
                # Start a print job
                job = win32print.StartDocPrinter(hprinter, 1, ("PRN Label", None, "RAW"))
                win32print.StartPagePrinter(hprinter)
                
                # Write the PRN content
                win32print.WritePrinter(hprinter, prn_content.encode('latin-1'))
                
                # End the print job
                win32print.EndPagePrinter(hprinter)
                win32print.EndDocPrinter(hprinter)
                
            finally:
                # Always close the printer
                win32print.ClosePrinter(hprinter)
                
            self.logger.info(f"Successfully printed to {printer_name}")
            return True
        except Exception as e:
            error_msg = f"Print error on {printer_name}: {str(e)}"
            self.show_status(error_msg, True)
            self.logger.error(error_msg)
            return False

    def process_and_print(self):
        try:
            # Read template file
            with open(self.prn_path.get(), 'rb') as f:
                template_content = f.read().decode('latin-1')
            
            self.logger.info(f"Loaded PRN template: {self.prn_path.get()}")

            # Read and validate CSV
            try:
                df = pd.read_csv(self.csv_path.get())
                required_columns = ['AFMPN', 'Country', 'QTY']
                missing_columns = [col for col in required_columns if col not in df.columns]
                if missing_columns:
                    raise ValueError(f"Missing required columns: {', '.join(missing_columns)}")
                
                self.logger.info(f"Loaded CSV file with {len(df)} rows: {self.csv_path.get()}")
            except Exception as e:
                error_msg = f"Error reading CSV file: {str(e)}"
                self.show_status(error_msg, True)
                self.logger.error(error_msg)
                self.toggle_controls(False)
                return

            # Calculate total labels
            total_to_print = df['QTY'].astype(int).sum()
            total_labels_printed = 0
            selected_printer = self.printer_combo.get()
            
            self.logger.info(f"Starting print job to {selected_printer}, total labels to print: {total_to_print}")
            now = datetime.now()
            # Process each row
            for index, row in df.iterrows():
                if not self.is_printing:
                    break

                try:
                    # Replace template variables
                    current_content = template_content
                    current_content = current_content.replace('@AFMPN@', str(row['AFMPN']))
                    current_content = current_content.replace('@CTT@', str(row['CTT']))
                    current_content = current_content.replace('@AFMPN5@', str(row['AFMPN'])[-5:])
                    current_content = current_content.replace('@Country@', str(row['Country']))
                    current_content = current_content.replace('@CountryCode@', self.CountrieCode_map[str(row['Country'])])
                    current_content = current_content.replace('@Date@', self.get_date_code())
                    current_content = current_content.replace("@mon@", now.strftime("%m"))  # Month (01-12)
                    current_content = current_content.replace("@day@", now.strftime("%d"))
                    current_content = current_content.replace("@yea@", now.strftime("%y"))  # Day of month (01-31)
                    current_content = current_content.replace("@hou@", now.strftime("%H"))  # Hour (00-23)
                    current_content = current_content.replace("@min@", now.strftime("%M"))  # Minute (00-59)
                    current_content = current_content.replace("@sec@", now.strftime("%S"))

                    qty = int(row['QTY'])
                    
                    self.logger.info(f"Processing design {index + 1}: {row['AFMPN']}, Quantity: {qty}")

                    # Print copies
                    for copy_num in range(1, qty + 1):
                        if not self.is_printing:
                            break

                        try:
                            # Print directly to the selected printer
                            success = self.print_to_printer(current_content, selected_printer)
                            
                            if success:
                                total_labels_printed += 1
                                self.update_progress(
                                    current_label=str(row['AFMPN']),
                                    copy_num=copy_num,
                                    total_copies=qty,
                                    design_num=index + 1,
                                    total_designs=len(df),
                                    total_labels_printed=total_labels_printed,
                                    total_to_print=total_to_print
                                )
                            else:
                                self.logger.error(f"Failed to print label {row['AFMPN']}, copy {copy_num}")
                                continue
                                
                        except Exception as e:
                            error_msg = f"Error printing label {row['AFMPN']}: {str(e)}"
                            self.show_status(error_msg, True)
                            self.logger.error(error_msg)
                            continue

                except Exception as e:
                    error_msg = f"Error processing row {index + 1}: {str(e)}"
                    self.show_status(error_msg, True)
                    self.logger.error(error_msg)
                    continue

            # Final status update
            if self.is_printing:
                completion_msg = f"Completed: Printed {total_labels_printed} labels from {len(df)} designs"
                self.show_status(completion_msg)
                self.logger.info(completion_msg)

        except Exception as e:
            error_msg = f"Error: {str(e)}"
            self.show_status(error_msg, True)
            self.logger.error(error_msg)
        finally:
            self.is_printing = False
            self.toggle_controls(False)


if __name__ == "__main__":
    app = PRNEditor()
    app.root.mainloop()