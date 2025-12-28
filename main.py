import dearpygui.dearpygui as dpg
import pandas as pd
import smtplib
import ssl
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import threading
import time
import json
import sys
import os
from datetime import datetime
import re
from pathlib import Path


class EmailAutomationApp:
    def __init__(self):
        self.contacts_df = None
        self.csv_columns = []  # Store all CSV columns for variable support
        self.email_subject = "Collaboration opportunity with ((CompanyName))"
        self.email_template = """Hey ((Name)),

I was impressed by your work at ((CompanyName)) and would love to explore potential collaboration opportunities.

I believe we could create something remarkable together.

Would you be available for a brief call next week?

Best regards,
((YourName))"""
        
        self.smtp_settings = {
            "smtp_server": "smtp.yourserver.com",
            "smtp_port": 465,
            "sender_email": "user@example.com",
            "sender_password": "password",
            "use_tls": True,
            "use_ssl": True  # Added this for port 465
        }
        self.sending_settings = {
            "delay_between_emails": 5,
            "max_emails_per_batch": 50,
            "your_name": "King",
            "bracket_style": "((double brackets))"
        }
        self.is_sending = False
        self.send_progress = 0
        self.total_emails = 0
        self.sent_emails = 0
        self.current_sending_thread = None
        self.sent_log = []
        
        self.window_width = 1500
        self.window_height = 1000
        
        # Load settings if they exist
        self.load_settings()
       
        # Initialize Dear PyGui
        dpg.create_context()

        font_size = 16
        header_font_size = 20
        font_path = "C:\\Windows\\Fonts\\calibri.ttf"
            
        with dpg.font_registry():
            with dpg.font(font_path, 20) as self.custom_font:
                # Load the Cyrillic character range
                dpg.add_font_range_hint(dpg.mvFontRangeHint_Cyrillic)

        self.setup_theme()
    
    
    def setup_theme(self):
        """Setup modern theme"""
        with dpg.theme() as global_theme:
            with dpg.theme_component(dpg.mvAll):
                # Modern dark theme
                dpg.add_theme_color(dpg.mvThemeCol_WindowBg, (18, 18, 24, 255))
                dpg.add_theme_color(dpg.mvThemeCol_ChildBg, (25, 25, 32, 255))
                dpg.add_theme_color(dpg.mvThemeCol_FrameBg, (40, 40, 48, 255))
                dpg.add_theme_color(dpg.mvThemeCol_FrameBgHovered, (50, 50, 60, 255))
                dpg.add_theme_color(dpg.mvThemeCol_FrameBgActive, (60, 60, 70, 255))
                dpg.add_theme_color(dpg.mvThemeCol_Button, (30, 100, 220, 255))
                dpg.add_theme_color(dpg.mvThemeCol_ButtonHovered, (40, 120, 240, 255))
                dpg.add_theme_color(dpg.mvThemeCol_ButtonActive, (20, 80, 200, 255))
                dpg.add_theme_color(dpg.mvThemeCol_Header, (30, 100, 220, 100))
                dpg.add_theme_color(dpg.mvThemeCol_HeaderHovered, (30, 100, 220, 150))
                dpg.add_theme_color(dpg.mvThemeCol_HeaderActive, (30, 100, 220, 200))
                dpg.add_theme_color(dpg.mvThemeCol_TitleBgActive, (30, 100, 220, 255))
                dpg.add_theme_color(dpg.mvThemeCol_TitleBg, (20, 20, 30, 255))
                dpg.add_theme_color(dpg.mvThemeCol_Text, (220, 220, 230, 255))
                dpg.add_theme_color(dpg.mvThemeCol_Border, (60, 60, 70, 255))
                dpg.add_theme_color(dpg.mvThemeCol_Separator, (60, 60, 70, 255))
                dpg.add_theme_color(dpg.mvThemeCol_Tab, (40, 40, 50, 255))
                dpg.add_theme_color(dpg.mvThemeCol_TabHovered, (50, 50, 65, 255))
                dpg.add_theme_color(dpg.mvThemeCol_TabActive, (30, 100, 220, 255))
                dpg.add_theme_color(dpg.mvThemeCol_TabUnfocusedActive, (40, 40, 50, 255))
                dpg.add_theme_color(dpg.mvThemeCol_MenuBarBg, (25, 25, 32, 255))
                
                # Modern rounded corners
                dpg.add_theme_style(dpg.mvStyleVar_FrameRounding, 4)
                dpg.add_theme_style(dpg.mvStyleVar_GrabRounding, 4)
                dpg.add_theme_style(dpg.mvStyleVar_TabRounding, 4)
                dpg.add_theme_style(dpg.mvStyleVar_ChildRounding, 8)
                dpg.add_theme_style(dpg.mvStyleVar_PopupRounding, 8)
                dpg.add_theme_style(dpg.mvStyleVar_WindowRounding, 8)
                
                # Spacing
                dpg.add_theme_style(dpg.mvStyleVar_ItemSpacing, 8, 6)
                dpg.add_theme_style(dpg.mvStyleVar_FramePadding, 10, 6)
                dpg.add_theme_style(dpg.mvStyleVar_CellPadding, 6, 4)
        
        dpg.bind_theme(global_theme)
    
    def create_windows(self):
        """Create windows with responsive design"""
        # Main window
        with dpg.window(label="Sales Email Automator", tag="Primary Window", 
                       width=self.window_width, height=self.window_height):
            dpg.bind_font(self.custom_font)
            
            # Menu bar
            with dpg.menu_bar():
                with dpg.menu(label="File"):
                    dpg.add_menu_item(label="Upload CSV Contacts", callback=self.open_csv_dialog)
                    dpg.add_menu_item(label="Clear CSV Contacts", callback=self.clear_contacts)
                    dpg.add_menu_item(label="Edit Settings", callback=lambda: dpg.show_item("SettingsWindow"))
                    dpg.add_menu_item(label="Save Settings", callback=self.save_settings)
                    dpg.add_separator()
                    dpg.add_menu_item(label="Exit", callback=lambda: dpg.stop_dearpygui())
                
                with dpg.menu(label="Tools"):
                    dpg.add_menu_item(label="Generate Sample CSV", callback=self.generate_sample_csv)

                # ADD THIS NEW "View" MENU:
                with dpg.menu(label="View"):
                    dpg.add_menu_item(label="Uploaded Contacts", 
                                     callback=lambda: self.show_contacts_popup())
                    dpg.add_menu_item(label="Sent Log", callback=lambda: dpg.show_item("LogWindow"))
            
            # Main responsive layout
            with dpg.group(horizontal=True, tag="MainLayout"):
                with dpg.child_window(tag="EditorPanel"):
                    self.create_editor_panel()
        
        # Settings window
        self.create_settings_window()
        
        # Log window
        self.create_log_window()

        # Create contacts popup window
        self.create_contacts_popup()

    def show_contacts_popup(self):
        """Show the uploaded contacts in a separate window"""
        # Recreate the popup window to refresh the data
        self.create_contacts_popup()
        dpg.show_item("ContactsPopupWindow")

    def clear_contacts(self):
        """Clear all stored contacts"""
        if self.contacts_df is not None and len(self.contacts_df) > 0:
            # Ask for confirmation
            if dpg.does_item_exist("ConfirmClearWindow"):
                dpg.delete_item("ConfirmClearWindow")
            
            with dpg.window(label="Confirm Clear", tag="ConfirmClearWindow", 
                           width=400, height=200, modal=True):
                dpg.bind_font(self.custom_font)
                dpg.add_spacer(height=20)
                dpg.add_text(f"Clear all {len(self.contacts_df)} contacts?")
                dpg.add_spacer(height=20)
                
                with dpg.group(horizontal=True):
                    dpg.add_button(label="Yes, Clear All", 
                                 callback=lambda: self._perform_clear_contacts(),
                                 width=150, height=35)
                    dpg.add_spacer(width=20)
                    dpg.add_button(label="Cancel", 
                                 callback=lambda: dpg.delete_item("ConfirmClearWindow"),
                                 width=150, height=35)
        else:
            self.update_status("No contacts to clear", "info")

    def _perform_clear_contacts(self):
        """Actually clear the contacts"""
        self.contacts_df = None
        self.csv_columns = []
        self.update_status("All contacts cleared", "success")
        dpg.delete_item("ConfirmClearWindow")
        
        # Also clear the contacts popup table if it exists
        if dpg.does_item_exist("ContactsPopupTable"):
            dpg.delete_item("ContactsPopupTable", children_only=True)
    
    def create_contacts_popup(self):
        """Create a separate window to show uploaded contacts"""
        if dpg.does_item_exist("ContactsPopupWindow"):
            dpg.delete_item("ContactsPopupWindow")
        
        with dpg.window(label="Uploaded Contacts", tag="ContactsPopupWindow", 
                       width=1000, height=800, show=False,  # Increased size
                       pos=[200, 150], on_close=lambda: dpg.hide_item("ContactsPopupWindow")):
            
            dpg.bind_font(self.custom_font)
            dpg.add_text("Uploaded Contacts", color=(30, 180, 255))
            dpg.add_separator()
            
            if self.contacts_df is None or len(self.contacts_df) == 0:
                dpg.add_text("No contacts loaded", color=(150, 150, 150))
            else:
                # Display contact count
                contact_count = len(self.contacts_df)
                dpg.add_text(f"Total contacts: {contact_count}", color=(100, 200, 100))
                dpg.add_spacer(height=10)
                
                # Create a scrollable group for the table
                with dpg.child_window(height=600, width=-1, horizontal_scrollbar=True):
                    # Create a table to display all contacts
                    with dpg.table(
                        tag="ContactsPopupTable",
                        header_row=True, 
                        borders_innerH=True, borders_outerH=True,
                        borders_innerV=True, borders_outerV=True,
                        reorderable=True, resizable=True,
                        scrollY=True, scrollX=True,  # Enable both scrollbars
                        height=580,
                        policy=dpg.mvTable_SizingFixedFit  # Better column sizing
                    ):
                        
                        # Add columns based on CSV structure
                        if len(self.csv_columns) > 0:
                            for col in self.csv_columns:
                                # Set initial width based on column name length
                                # Minimum width of 100, more for longer column names
                                initial_width = max(100, len(str(col)) * 8)
                                dpg.add_table_column(
                                    label=col, 
                                    width=initial_width,
                                    width_stretch=False  # Don't stretch, allow manual resizing
                                )
                        else:
                            # Default columns if no CSV loaded yet
                            for col in ['Name', 'Company', 'Email', 'Position']:
                                dpg.add_table_column(
                                    label=col, 
                                    width=150,
                                    width_stretch=False
                                )
                        
                        # Add data rows (limit for performance)
                        if self.contacts_df is not None:
                            display_limit = min(500, contact_count)  # Limit for performance
                            
                            for i in range(display_limit):
                                with dpg.table_row():
                                    for col in self.csv_columns:
                                        value = str(self.contacts_df.iloc[i].get(col, ''))
                                        if pd.isna(value):
                                            value = ''
                                        # Truncate very long values for performance
                                        if len(value) > 200:
                                            value = value[:197] + "..."
                                        dpg.add_text(value)
                            
                            if contact_count > display_limit:
                                dpg.add_spacer(height=5)
                                dpg.add_text(f"Showing first {display_limit} of {contact_count} contacts", 
                                           color=(150, 150, 150))
    
    def create_editor_panel(self):
        """Create the email editor panel with responsive design"""
        with dpg.group():
            dpg.add_text("Email Template Editor", color=(30, 180, 255))
            dpg.add_separator()
            
            # Email subject
            dpg.add_text("Subject:")
            dpg.add_input_text(tag="EmailSubject", 
                              default_value=self.email_subject, 
                               callback=None,
                               width=-1)
            
            dpg.add_spacer(height=10)
            
            # Email template editor
            dpg.add_text("Body:")
            editor_height = int(self.window_height * 0.45)
            dpg.add_input_text(tag="EmailTemplate", multiline=True, 
                              default_value=self.email_template,
                              width=-1, height=editor_height,
                               callback=None, 
                              tab_input=True)
            
            dpg.add_spacer(height=15)
            
            # Preview and send section
            with dpg.group(horizontal=True):
                dpg.add_button(label="Preview Email", callback=self.preview_email, 
                              width=150, height=35)
                dpg.add_spacer(width=20)
                dpg.add_button(label="Send Emails", callback=self.start_sending_emails, 
                              width=150, height=35, tag="SendButton")
                dpg.add_button(label="Stop Sending", callback=self.stop_sending_emails,
                              width=150, height=35, show=False, tag="StopButton")
            
            dpg.add_spacer(height=15)
            
            # Progress section
            dpg.add_text("Sending Progress", color=(30, 180, 255))
            dpg.add_separator()
            
            dpg.add_progress_bar(tag="SendProgress", default_value=0, width=-1, height=20)
            
            dpg.add_text("Ready to send", tag="StatusText", color=(150, 220, 150))
    
    def create_settings_window(self):
        """Create settings window"""
        with dpg.window(label="Settings", tag="SettingsWindow", 
                       width=700,height=550, show=False, 
                       pos=[200, 200], on_close=lambda: dpg.hide_item("SettingsWindow")):
            
            dpg.bind_font(self.custom_font)
            with dpg.tab_bar():
                # SMTP Settings tab
                with dpg.tab(label="Email Server"):
                    self.create_smtp_settings()
                
                # Sending Settings tab
                with dpg.tab(label="Sending Settings"):
                    self.create_sending_settings()
    
    def create_smtp_settings(self):
        """Create SMTP settings section"""
        dpg.add_text("SMTP Configuration", color=(30, 180, 255))
        dpg.add_separator()
        
        dpg.add_text("Server:")
        dpg.add_input_text(tag="SMTPServer", 
                          default_value=self.smtp_settings["smtp_server"], 
                          width=-1)
        
        dpg.add_text("Port:")
        dpg.add_input_int(tag="SMTPPort", 
                         default_value=self.smtp_settings["smtp_port"], 
                         width=-1, min_value=1, max_value=65535)
        
        dpg.add_text("Email Address:")
        dpg.add_input_text(tag="SenderEmail", 
                          default_value=self.smtp_settings["sender_email"], 
                          width=-1)
        
        dpg.add_text("Password/App Password:")
        dpg.add_input_text(tag="SenderPassword", password=True,
                          default_value=self.smtp_settings["sender_password"], 
                          width=-1)
        
        dpg.add_checkbox(tag="UseTLS", label="Use TLS/SSL", 
                        default_value=self.smtp_settings["use_tls"])
        
        dpg.add_spacer(height=20)
        dpg.add_button(label="Test Connection", 
                      callback=self.test_smtp_connection, 
                      width=200, height=35)
        
        dpg.add_spacer(height=10)
        dpg.add_text("For Beget.com, port 465 with SSL is typically used.", 
                    color=(255, 180, 50))

    def update_bracket_style(self, sender, app_data):
        """Update bracket style immediately when changed in dropdown"""
        # Update the setting immediately
        self.sending_settings["bracket_style"] = app_data
        print(f"Bracket style updated to: {app_data}")  # Debug print
    
    def create_sending_settings(self):
        """Create sending settings section"""
        dpg.add_text("Sending Configuration", color=(30, 180, 255))
        dpg.add_separator()
        
        dpg.add_text("Your Name (used in ((YourName)):")
        dpg.add_input_text(tag="YourName", 
                          default_value=self.sending_settings["your_name"], 
                          width=-1)
        
        dpg.add_text("Delay between emails (seconds):")
        dpg.add_input_int(tag="EmailDelay", 
                         default_value=self.sending_settings["delay_between_emails"], 
                         width=-1, min_value=1, max_value=3600)
        
        dpg.add_text("Maximum emails per batch:")
        dpg.add_input_int(tag="MaxEmails", 
                         default_value=self.sending_settings["max_emails_per_batch"], 
                         width=-1, min_value=1, max_value=1000)

        dpg.add_spacer(height=20)
        dpg.add_text("Variable Bracket Style:", color=(30, 180, 255))
        dpg.add_separator()
        dpg.add_text("Choose how variables appear in your template:")
        bracket_options = [
            "(single brackets)",
            "((double brackets))",
            "(((triple brackets)))",
            "{curly brackets}",
            "[square brackets]",
            "{{double curly brackets}}",
            "[[double square brackets]]"
        ]

        # Find the current bracket style index
        current_style = self.sending_settings.get("bracket_style", "((double brackets))")
        current_index = bracket_options.index(current_style) if current_style in bracket_options else 1  # Default to double brackets
        
        # Create the dropdown
        dpg.add_combo(
            tag="BracketStyle",
            items=bracket_options,
            default_value=bracket_options[current_index],
            width=-1,
            callback=self.update_bracket_style
        )
        
        dpg.add_spacer(height=20)
        dpg.add_text("Be respectful with sending rates to avoid being flagged as spam.", 
                    color=(255, 150, 50))
    
    def create_log_window(self):
        """Create log window"""
        with dpg.window(label="Email Log", tag="LogWindow", 
                       width=900, height=600, show=False,
                       pos=[250, 250], on_close=lambda: dpg.hide_item("LogWindow")):
            
            dpg.bind_font(self.custom_font)
            dpg.add_text("Sent Email History", color=(30, 180, 255))
            dpg.add_separator()
            
            with dpg.table(tag="LogTable", header_row=True,
                          borders_innerH=True, borders_outerH=True,
                          borders_innerV=True, borders_outerV=True,
                          resizable=True, reorderable=True,
                          scrollY=True, height=500):
                dpg.add_table_column(label="Time", width_fixed=True, width=150)
                dpg.add_table_column(label="Recipient", width_stretch=True)
                dpg.add_table_column(label="Company", width_stretch=True)
                dpg.add_table_column(label="Status", width_fixed=True, width=120)
    
    def open_csv_dialog(self):
        """Open file dialog for CSV selection"""
        dpg.show_item("file_dialog_id")
    
    def load_csv_file(self, sender, app_data):
        """Load and process CSV file - appends to existing contacts"""
        try:
            print(f"Debug - app_data keys: {app_data.keys() if hasattr(app_data, 'keys') else 'no keys'}")
            print(f"Debug - app_data: {app_data}")
            
            # Different Dear PyGui versions have different data structures
            if "file_path_name" in app_data:
                file_path = app_data["file_path_name"]
            elif "selections" in app_data and app_data["selections"]:
                # In some versions, selections is a dict
                selections = app_data["selections"]
                if isinstance(selections, dict):
                    file_path = list(selections.values())[0]
                else:
                    file_path = selections
            elif "current_path" in app_data:
                file_path = app_data["current_path"]
            else:
                # Try to get the first value if it's a simple structure
                try:
                    if isinstance(app_data, (list, tuple)) and len(app_data) > 0:
                        file_path = app_data[0]
                    elif isinstance(app_data, str):
                        file_path = app_data
                    else:
                        self.update_status("No file path found in selection", "error")
                        return
                except:
                    self.update_status("Could not parse file selection", "error")
                    return
            
            print(f"Loading file: {file_path}")
            
            # Fix the file extension issue - check if it has .* and try to find .csv
            if file_path.endswith('.*'):
                # Try to find a .csv version
                csv_path = file_path[:-2] + '.csv'
                if os.path.exists(csv_path):
                    file_path = csv_path
                    print(f"Fixed file path to: {file_path}")
                else:
                    self.update_status(f"File not found: {csv_path}", "error")
                    return
            
            # Check if file exists
            if not os.path.exists(file_path):
                self.update_status(f"File not found: {file_path}", "error")
                return
            
            # Store current contacts to append to
            current_df = self.contacts_df.copy() if self.contacts_df is not None else None
            current_columns = self.csv_columns.copy() if self.csv_columns else []
            
            # First try with semicolon delimiter (the original reliable method)
            print("Trying semicolon delimiter with UTF-8...")
            new_df = None
            delimiter_used = None
            encoding_used = None
            
            try:
                new_df = pd.read_csv(file_path, delimiter=';', encoding='utf-8')
                delimiter_used = ';'
                encoding_used = 'utf-8'
                print("Success with semicolon delimiter")
            except Exception as e1:
                print(f"Semicolon UTF-8 failed: {e1}")
                
                # Try semicolon with other encodings
                try:
                    new_df = pd.read_csv(file_path, delimiter=';', encoding='utf-8-sig')
                    delimiter_used = ';'
                    encoding_used = 'utf-8-sig'
                    print("Success with semicolon delimiter and utf-8-sig")
                except Exception as e2:
                    print(f"Semicolon utf-8-sig failed: {e2}")
                    try:
                        new_df = pd.read_csv(file_path, delimiter=';', encoding='windows-1251')
                        delimiter_used = ';'
                        encoding_used = 'windows-1251'
                        print("Success with semicolon delimiter and windows-1251")
                    except Exception as e3:
                        print(f"Semicolon windows-1251 failed: {e3}")
                        
                        # If semicolon fails, try comma delimiter
                        print("Trying comma delimiter...")
                        try:
                            new_df = pd.read_csv(file_path, delimiter=',', encoding='utf-8')
                            delimiter_used = ','
                            encoding_used = 'utf-8'
                            print("Success with comma delimiter")
                        except Exception as e4:
                            print(f"Comma UTF-8 failed: {e4}")
                            try:
                                new_df = pd.read_csv(file_path, delimiter=',', encoding='utf-8-sig')
                                delimiter_used = ','
                                encoding_used = 'utf-8-sig'
                                print("Success with comma delimiter and utf-8-sig")
                            except Exception as e5:
                                print(f"Comma utf-8-sig failed: {e5}")
                                try:
                                    new_df = pd.read_csv(file_path, delimiter=',', encoding='windows-1251')
                                    delimiter_used = ','
                                    encoding_used = 'windows-1251'
                                    print("Success with comma delimiter and windows-1251")
                                except Exception as e6:
                                    print(f"Comma windows-1251 failed: {e6}")
                                    
                                    # Last resort: try with pandas auto-detection (no delimiter specified)
                                    try:
                                        new_df = pd.read_csv(file_path, encoding='utf-8')
                                        delimiter_used = 'auto'
                                        encoding_used = 'utf-8'
                                        print("Success with auto-detection")
                                    except Exception as e7:
                                        print(f"Auto-detection failed: {e7}")
                                        # Try with error handling
                                        try:
                                            new_df = pd.read_csv(
                                                file_path, 
                                                encoding='utf-8',
                                                on_bad_lines='skip',
                                                engine='python'
                                            )
                                            delimiter_used = 'auto'
                                            encoding_used = 'utf-8'
                                            print("Success with error handling")
                                        except Exception as e8:
                                            print(f"All attempts failed: {e8}")
                                            self.update_status(f"Error loading CSV: Could not parse file with any method", "error")
                                            return
            
            if new_df is None or len(new_df.columns) == 0:
                self.update_status("Error: No valid data found in CSV file", "error")
                return
            
            # Clean column names for the new dataframe
            new_columns = [str(col).strip() for col in new_df.columns.tolist()]
            new_df.columns = new_columns
            
            # If we only have 1 column and delimiter was comma, maybe it's actually semicolon
            # or vice versa - let's do a simple check
            if len(new_columns) == 1:
                print(f"Warning: Only 1 column detected with delimiter '{delimiter_used}'")
                
                # Check the first few rows to see what delimiter might be present
                try:
                    with open(file_path, 'r', encoding=encoding_used, errors='ignore') as f:
                        first_lines = [f.readline() for _ in range(3)]
                    
                    for line in first_lines:
                        # Count different delimiters
                        semicolon_count = line.count(';')
                        comma_count = line.count(',')
                        
                        print(f"Line sample: {line[:100]}...")
                        print(f"Semicolons: {semicolon_count}, Commas: {comma_count}")
                        
                        # If there are semicolons but we used comma delimiter
                        if semicolon_count > 0 and delimiter_used == ',':
                            print("Trying semicolon instead...")
                            try:
                                new_df = pd.read_csv(file_path, delimiter=';', encoding=encoding_used)
                                delimiter_used = ';'
                                new_columns = [str(col).strip() for col in new_df.columns.tolist()]
                                new_df.columns = new_columns
                                print("Fixed: Now using semicolon delimiter")
                            except:
                                pass
                        # If there are commas but we used semicolon delimiter
                        elif comma_count > 0 and delimiter_used == ';':
                            print("Trying comma instead...")
                            try:
                                new_df = pd.read_csv(file_path, delimiter=',', encoding=encoding_used)
                                delimiter_used = ','
                                new_columns = [str(col).strip() for col in new_df.columns.tolist()]
                                new_df.columns = new_columns
                                print("Fixed: Now using comma delimiter")
                            except:
                                pass
                        
                        if len(new_columns) > 1:
                            break
                            
                except Exception as e:
                    print(f"Error checking delimiters: {e}")
            
            # Now append to existing contacts
            if current_df is not None and len(current_df) > 0:
                # Check if column structures match
                if set(current_columns) != set(new_columns):
                    # Columns don't match - ask user what to do
                    self.handle_column_mismatch(current_df, new_df, current_columns, new_columns, file_path)
                    return
                
                # Append the data
                combined_df = pd.concat([current_df, new_df], ignore_index=True)
                self.contacts_df = combined_df
                self.csv_columns = current_columns  # Keep original column order
                
                # Remove duplicate emails if any
                if 'Email' in self.contacts_df.columns:
                    before_dedup = len(self.contacts_df)
                    self.contacts_df = self.contacts_df.drop_duplicates(subset=['Email'], keep='first')
                    after_dedup = len(self.contacts_df)
                    duplicates_removed = before_dedup - after_dedup
                    
                    new_count = len(new_df)
                    total_count = len(self.contacts_df)
                    
                    status_msg = f"Added {new_count} contacts from {os.path.basename(file_path)}. "
                    status_msg += f"Total: {total_count} contacts. "
                    if duplicates_removed > 0:
                        status_msg += f"Removed {duplicates_removed} duplicate emails."
                    
                    self.update_status(status_msg, "success")
                else:
                    new_count = len(new_df)
                    total_count = len(self.contacts_df)
                    self.update_status(f"Added {new_count} contacts from {os.path.basename(file_path)}. Total: {total_count} contacts", "success")
            else:
                # First time loading or no current contacts
                self.contacts_df = new_df
                self.csv_columns = new_columns
                
                new_count = len(new_df)
                self.update_status(f"Loaded {new_count} contacts from {os.path.basename(file_path)}", "success")
            
            # Update the contacts popup window to show all contacts
            if dpg.does_item_exist("ContactsPopupWindow"):
                self.create_contacts_popup()
            
            # Debug: print column names
            print(f"Columns loaded: {self.csv_columns}")
            print(f"Total contacts: {len(self.contacts_df)}")
            print(f"Delimiter used: {delimiter_used}")
            print(f"Encoding used: {encoding_used}")
            
        except Exception as e:
            error_msg = f"Error loading CSV: {str(e)}"
            print(f"Full error: {error_msg}")
            print(f"Error type: {type(e)}")
            import traceback
            traceback.print_exc()
            self.update_status(error_msg, "error")

    def handle_column_mismatch(self, current_df, new_df, current_columns, new_columns, file_path):
        """Handle case where new CSV has different columns than existing contacts"""
        if dpg.does_item_exist("ColumnMismatchWindow"):
            dpg.delete_item("ColumnMismatchWindow")
        
        with dpg.window(label="Column Mismatch", tag="ColumnMismatchWindow", 
                       width=600, height=400, modal=True):
            dpg.bind_font(self.custom_font)
            dpg.add_spacer(height=10)
            dpg.add_text("The new CSV file has different columns than existing contacts:", color=(255, 180, 50))
            dpg.add_separator()
            
            dpg.add_text("Existing columns:", color=(100, 200, 100))
            existing_cols_text = ", ".join(current_columns)
            dpg.add_text(existing_cols_text)
            
            dpg.add_spacer(height=10)
            
            dpg.add_text("New file columns:", color=(100, 200, 100))
            new_cols_text = ", ".join(new_columns)
            dpg.add_text(new_cols_text)
            
            dpg.add_spacer(height=20)
            dpg.add_text("How do you want to proceed?", color=(200, 200, 100))
            
            with dpg.group(horizontal=True):
                dpg.add_button(label="Replace all contacts", 
                             callback=lambda: self._replace_all_contacts(new_df, new_columns, file_path),
                             width=200, height=40)
                dpg.add_spacer(width=20)
                dpg.add_button(label="Keep only matching columns", 
                             callback=lambda: self._keep_matching_columns(new_df, current_columns, new_columns, file_path),
                             width=300, height=40)
                dpg.add_spacer(width=20)
                dpg.add_button(label="Cancel import", 
                             callback=lambda: dpg.delete_item("ColumnMismatchWindow"),
                             width=150, height=40)

    def _replace_all_contacts(self, new_df, new_columns, file_path):
        """Replace all existing contacts with new ones"""
        self.contacts_df = new_df
        self.csv_columns = new_columns
        
        new_count = len(new_df)
        self.update_status(f"Replaced all contacts with {new_count} contacts from {os.path.basename(file_path)}", "warning")
        dpg.delete_item("ColumnMismatchWindow")
        
        # Update the contacts popup window
        if dpg.does_item_exist("ContactsPopupWindow"):
            self.create_contacts_popup()

    def _keep_matching_columns(self, new_df, current_columns, new_columns, file_path):
        """Keep only columns that exist in both datasets"""
        # Find common columns
        common_columns = [col for col in current_columns if col in new_columns]
        
        if not common_columns:
            self.update_status("Error: No common columns found between datasets", "error")
            dpg.delete_item("ColumnMismatchWindow")
            return
        
        # Filter new dataframe to only keep common columns
        filtered_new_df = new_df[[col for col in common_columns if col in new_df.columns]].copy()
        
        # Append to existing dataframe (which already has all common columns)
        combined_df = pd.concat([self.contacts_df[common_columns], filtered_new_df], ignore_index=True)
        self.contacts_df = combined_df
        self.csv_columns = common_columns
        
        new_count = len(filtered_new_df)
        total_count = len(self.contacts_df)
        
        status_msg = f"Added {new_count} contacts (keeping {len(common_columns)} common columns). "
        status_msg += f"Total: {total_count} contacts."
        
        self.update_status(status_msg, "success")
        dpg.delete_item("ColumnMismatchWindow")
        
        # Update the contacts popup window
        if dpg.does_item_exist("ContactsPopupWindow"):
            self.create_contacts_popup()
        
    def update_table_columns(self):
        """Update table columns based on CSV structure"""
        # Clear existing columns
        if dpg.does_item_exist("ContactsTable"):
            dpg.delete_item("ContactsTable", children_only=True)
        
        # Add columns for important fields
        important_columns = ['Name', 'CompanyName', 'Email', 'Position']
        
        for col in important_columns:
            if col in self.csv_columns:
                dpg.add_table_column(label=col, width_stretch=True, parent="ContactsTable")
    
    def insert_variable(self, variable):
        """Insert a variable at cursor position in email template"""
        current_text = dpg.get_value("EmailTemplate")
        dpg.set_value("EmailTemplate", current_text + f"[{variable}]")
    
    def replace_variables(self, text, contact):
        """Replace ALL template variables with contact data using selected bracket style"""
        # Get the current bracket style
        bracket_style = self.sending_settings.get("bracket_style", "((double brackets))")
        
        # Map bracket styles to actual bracket characters
        bracket_map = {
            "(single brackets)": ("(", ")"),
            "((double brackets))": ("((", "))"),
            "(((triple brackets)))": ("(((", ")))"),
            "{curly brackets}": ("{", "}"),
            "[square brackets]": ("[", "]"),
            "{{double curly brackets}}": ("{{", "}}"),
            "[[double square brackets]]": ("[[", "]]")
        }
        
        # Get the actual bracket characters for the selected style
        left_bracket, right_bracket = bracket_map.get(bracket_style, ("((", "))"))
        
        # Add YourName to contact data for replacement
        contact_with_yourname = contact.copy() if isinstance(contact, dict) else contact.to_dict()
        contact_with_yourname['YourName'] = self.sending_settings["your_name"]
        
        # Create a list of all variables to replace (including YourName)
        all_variables = list(self.csv_columns) + ['YourName']
        
        # Replace ALL variables using the selected bracket style
        for variable in all_variables:
            if variable in contact_with_yourname:
                # Create the variable pattern with the selected brackets
                variable_pattern = f'{left_bracket}{variable}{right_bracket}'
                value = str(contact_with_yourname[variable])
                if pd.isna(value):
                    value = ''
                # Replace ALL occurrences of this variable
                text = text.replace(variable_pattern, value)
        
        return text
    
    def generate_sample_csv(self):
        """Generate a sample CSV file for testing"""
        sample_data = {
            'Name': ['John', 'Maria', 'Alex', 'Sophie', 'Ivan'],
            'Surname': ['Smith', 'Garcia', 'Johnson', 'Dubois', 'Petrov'],
            'CompanyName': ['TechCorp', 'InnovateLabs', 'DataSystems', 'ParisTech', 'RusSoft'],
            'Email': ['john.smith@techcorp.com', 'maria.g@innovatelabs.com', 
                     'alex.j@datasystems.com', 'sophie.d@paristech.fr', 'ivan.p@russoft.ru'],
            'Sex': ['Male', 'Female', 'Male', 'Female', 'Male'],
            'Position': ['CEO', 'CTO', 'Marketing Director', 'Product Manager', 'Lead Developer'],
            'Industry': ['Technology', 'Biotech', 'Data Analytics', 'Software', 'Fintech'],
            'Location': ['New York', 'Boston', 'Chicago', 'Paris', 'Moscow'],
            'Phone': ['+1-555-0101', '+1-555-0102', '+1-555-0103', '+33-1-5555', '+7-495-1234'],
            'LastContact': ['2023-10-15', '2023-11-20', '2023-09-05', '2023-12-01', '2023-11-15']
        }
        
        df = pd.DataFrame(sample_data)
        
        # Save to current directory
        try:
            file_path = Path.cwd() / "sample_contacts.csv"
            df.to_csv(file_path, index=False)
            
            self.update_status(f"Sample CSV created: {file_path}", "success")
            self.show_message("Success", f"Sample CSV file created at:\n{file_path}")
        except Exception as e:
            self.update_status(f"Error creating sample CSV: {str(e)}", "error")
    
    def preview_email(self):
        # self.update_settings_from_ui()
        """Preview email with actual data"""
        if self.contacts_df is None or len(self.contacts_df) == 0:
            self.show_message("Error", "Please load contacts first")
            return
        
        # Get the first contact for preview
        contact = self.contacts_df.iloc[0]
        
        # Get template and subject
        template = dpg.get_value("EmailTemplate")
        subject = dpg.get_value("EmailSubject")
        
        # Replace variables
        preview_body = self.replace_variables(template, contact)
        preview_subject = self.replace_variables(subject, contact)
        
        # Show preview in a new window
        if dpg.does_item_exist("PreviewWindow"):
            dpg.delete_item("PreviewWindow")
        
        with dpg.window(label="Email Preview", tag="PreviewWindow", width=1200, height=700, modal=True):
            dpg.bind_font(self.custom_font)
            dpg.add_text("Subject:", color=(30, 180, 255))
            dpg.add_text(preview_subject)
            dpg.add_spacer(height=10)
            dpg.add_text("Body:", color=(30, 180, 255))
            dpg.add_input_text(tag="PreviewContent", multiline=True, 
                              default_value=preview_body, width=1150, height=600,
                              readonly=True)
    
    def start_sending_emails(self):
        """Start sending emails"""
        if self.contacts_df is None or len(self.contacts_df) == 0:
            self.show_message("Error", "Please load contacts first")
            return
        
        if not self.validate_smtp_settings():
            self.show_message("Error", "Please configure SMTP settings first")
            dpg.show_item("SettingsWindow")
            return
        
        # Disable send button, enable stop button
        dpg.hide_item("SendButton")
        dpg.show_item("StopButton")
        
        # Start sending in a separate thread
        self.is_sending = True
        self.sent_emails = 0
        self.total_emails = len(self.contacts_df)
        self.send_progress = 0
        
        self.current_sending_thread = threading.Thread(target=self.send_emails_thread, daemon=True)
        self.current_sending_thread.start()
    
    def send_emails_thread(self):
        """Background thread for sending emails"""
        try:
            # Get email content
            template = dpg.get_value("EmailTemplate")
            subject = dpg.get_value("EmailSubject")
            
            # Setup SMTP connection
            smtp_server = self.smtp_settings["smtp_server"]
            smtp_port = self.smtp_settings["smtp_port"]
            sender_email = self.smtp_settings["sender_email"]
            sender_password = self.smtp_settings["sender_password"]
            use_tls = self.smtp_settings["use_tls"]
            
            print(f"Connecting to {smtp_server}:{smtp_port}...")
            
            # For port 465 (SSL), use SMTP_SSL instead of SMTP with starttls
            if smtp_port == 465:
                print("Using SSL for port 465")
                context = ssl.create_default_context()
                server = smtplib.SMTP_SSL(smtp_server, smtp_port, context=context)
            elif use_tls:
                print("Using TLS")
                context = ssl.create_default_context()
                server = smtplib.SMTP(smtp_server, smtp_port)
                server.starttls(context=context)
            else:
                print("Using plain connection")
                server = smtplib.SMTP(smtp_server, smtp_port)
            
            print("Logging in...")
            server.login(sender_email, sender_password)
            print("Login successful")
            
            # Send emails
            for i, contact in self.contacts_df.iterrows():
                if not self.is_sending:
                    break
                
                try:
                    # Prepare email
                    recipient_email = str(contact.get('Email', '')) if 'Email' in contact else None
                    
                    if not recipient_email or not re.match(r"[^@]+@[^@]+\.[^@]+", recipient_email):
                        self.log_sent_email(contact, "Invalid Email")
                        continue
                    
                    # Create message
                    msg = MIMEMultipart()
                    msg['From'] = sender_email
                    msg['To'] = recipient_email
                    msg['Subject'] = self.replace_variables(subject, contact)
                    
                    # Create body
                    body = self.replace_variables(template, contact)
                    msg.attach(MIMEText(body, 'plain'))
                    
                    # Send email
                    print(f"Sending to {recipient_email}...")
                    server.send_message(msg)
                    
                    # Update progress
                    self.sent_emails += 1
                    self.send_progress = (self.sent_emails / self.total_emails) * 100
                    
                    # Log success
                    self.log_sent_email(contact, "Sent")
                    self.update_status(f"Sent to {contact.get('Name', 'N/A')}", "info")
                    
                    # Delay between emails
                    if i < len(self.contacts_df) - 1 and self.is_sending:
                        time.sleep(self.sending_settings["delay_between_emails"])
                        
                except Exception as e:
                    self.log_sent_email(contact, f"Error: {str(e)[:30]}")
                    self.update_status(f"Failed: {str(e)[:50]}", "error")
            
            # Close SMTP connection
            server.quit()
            
            # Update UI when done
            if self.is_sending:
                self.update_status(f"Sent {self.sent_emails} emails", "success")
            else:
                self.update_status(f"Stopped. Sent {self.sent_emails} emails", "warning")
                
        except Exception as e:
            error_msg = f"SMTP Error: {str(e)}"
            print(error_msg)
            self.update_status(error_msg, "error")
        
        # Re-enable send button
        self.is_sending = False
        dpg.show_item("SendButton")
        dpg.hide_item("StopButton")
    
    def stop_sending_emails(self):
        """Stop sending emails"""
        self.is_sending = False
        self.update_status("Stopping...", "warning")
    
    def log_sent_email(self, contact, status):
        """Log sent email to history"""
        log_entry = {
            "time": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "recipient": str(contact.get('Email', 'N/A')),
            "status": status,
            "company": str(contact.get('CompanyName', 'N/A'))
        }
        self.sent_log.append(log_entry)
        
        # Update log window if visible
        if dpg.does_item_exist("LogTable"):
            with dpg.table_row(parent="LogTable"):
                dpg.add_text(log_entry["time"])
                dpg.add_text(log_entry["recipient"])
                status_color = (0, 255, 0) if status == "Sent" else (255, 100, 0)
                dpg.add_text(log_entry["status"], color=status_color)
                dpg.add_text(log_entry["company"])
    
    def update_status(self, message, status_type="info"):
        """Update status text with color coding"""
        colors = {
            "success": (0, 255, 100),
            "error": (255, 50, 50),
            "warning": (255, 150, 0),
            "info": (150, 200, 255)
        }
        
        dpg.set_value("StatusText", message)
        dpg.configure_item("StatusText", color=colors.get(status_type, (150, 200, 255)))
        
        # Update progress bar
        if status_type == "success" and self.send_progress == 100:
            dpg.set_value("SendProgress", 100)
        elif self.is_sending:
            dpg.set_value("SendProgress", self.send_progress)
    
    def test_smtp_connection(self):
        """Test SMTP connection with current settings"""
        if not self.validate_smtp_settings():
            self.show_message("Error", "Please fill all SMTP fields")
            return
        
        try:
            # Get values from UI
            self.smtp_settings["smtp_server"] = dpg.get_value("SMTPServer")
            self.smtp_settings["smtp_port"] = dpg.get_value("SMTPPort")
            self.smtp_settings["sender_email"] = dpg.get_value("SenderEmail")
            self.smtp_settings["sender_password"] = dpg.get_value("SenderPassword")
            self.smtp_settings["use_tls"] = dpg.get_value("UseTLS")
            
            smtp_server = self.smtp_settings["smtp_server"]
            smtp_port = self.smtp_settings["smtp_port"]
            sender_email = self.smtp_settings["sender_email"]
            sender_password = self.smtp_settings["sender_password"]
            use_tls = self.smtp_settings["use_tls"]
            
            self.update_status("Testing SMTP connection...", "info")
            
            # For port 465 (SSL), use SMTP_SSL instead of SMTP with starttls
            if smtp_port == 465:
                context = ssl.create_default_context()
                server = smtplib.SMTP_SSL(smtp_server, smtp_port, context=context)
            elif use_tls:
                context = ssl.create_default_context()
                server = smtplib.SMTP(smtp_server, smtp_port)
                server.starttls(context=context)
            else:
                server = smtplib.SMTP(smtp_server, smtp_port)
            
            server.login(sender_email, sender_password)
            server.quit()
            
            self.show_message("Success", "SMTP connection test successful!")
            self.update_status("SMTP connection test successful", "success")
            
        except Exception as e:
            self.show_message("Error", f"SMTP connection failed: {str(e)}")
            self.update_status(f"SMTP test failed: {str(e)[:50]}", "error")
    
    def validate_smtp_settings(self):
        """Check if SMTP settings are filled"""
        return (self.smtp_settings["smtp_server"] and 
                self.smtp_settings["sender_email"] and 
                self.smtp_settings["sender_password"])
    
    def save_settings(self):
        """Save settings to JSON file"""
        try:
            # Update settings from UI if settings window exists
            if dpg.does_item_exist("SMTPServer"):
                self.smtp_settings["smtp_server"] = dpg.get_value("SMTPServer")
                self.smtp_settings["smtp_port"] = dpg.get_value("SMTPPort")
                self.smtp_settings["sender_email"] = dpg.get_value("SenderEmail")
                self.smtp_settings["sender_password"] = dpg.get_value("SenderPassword")
                self.smtp_settings["use_tls"] = dpg.get_value("UseTLS")
                
                self.sending_settings["your_name"] = dpg.get_value("YourName")
                self.sending_settings["delay_between_emails"] = dpg.get_value("EmailDelay")
                self.sending_settings["max_emails_per_batch"] = dpg.get_value("MaxEmails")

            # Save current email draft from UI
            if dpg.does_item_exist("EmailTemplate"):
                self.email_template = dpg.get_value("EmailTemplate")

            if dpg.does_item_exist("BracketStyle"):
                bracket_style = dpg.get_value("BracketStyle")
                self.sending_settings["bracket_style"] = bracket_style

            # Save email subject from UI
            if dpg.does_item_exist("EmailSubject"):
                self.email_subject = dpg.get_value("EmailSubject")
            
            # Save to file
            settings = {
                "smtp_settings": self.smtp_settings,
                "sending_settings": self.sending_settings,
                "email_template": self.email_template,
                "email_subject": self.email_subject
            }
            
            with open("email_automation_settings.json", "w") as f:
                json.dump(settings, f, indent=2)
            
            self.update_status("Settings saved successfully", "success")
        except Exception as e:
            self.update_status(f"Error saving settings: {str(e)}", "error")
    
    def load_settings(self):
        """Load settings from JSON file"""
        try:
            if os.path.exists("email_automation_settings.json"):
                with open("email_automation_settings.json", "r") as f:
                    settings = json.load(f)
                
                self.smtp_settings.update(settings.get("smtp_settings", {}))
                self.sending_settings.update(settings.get("sending_settings", {}))
                self.email_template = settings.get("email_template", self.email_template)
                self.email_subject = settings.get("email_subject", self.email_subject)
                
                print("Settings loaded from file")
        except Exception as e:
            print(f"Note: Could not load settings: {e}")
    
    def show_message(self, title, message):
        """Show a simple message box"""
        if dpg.does_item_exist("MessageWindow"):
            dpg.delete_item("MessageWindow")
        
        with dpg.window(label=title, tag="MessageWindow", width=400, height=200, modal=True):
            dpg.bind_font(self.custom_font)
            dpg.add_spacer(height=20)
            dpg.add_text(message)
            dpg.add_spacer(height=20)
            dpg.add_button(label="OK", width=100, callback=lambda: dpg.delete_item("MessageWindow"))
    
    def run(self):
        """Run the application"""
        # Create file dialog
        with dpg.file_dialog(
            directory_selector=False, 
            show=False, 
            callback=self.load_csv_file,
            tag="file_dialog_id", 
            width=900, 
            height=700,
            default_path=str(Path.cwd())
        ):
            dpg.add_file_extension(".csv", color=(0, 255, 100, 255))
            dpg.add_file_extension(".*", color=(150, 150, 150, 255))
        
        # Create windows
        self.create_windows()
        
        # Setup viewport
        dpg.create_viewport(
            title='Liberty Mail Automation', 
            width=self.window_width, 
            height=self.window_height,
            min_width=900,
            min_height=600
        )

        import ctypes
        try:
            ctypes.windll.user32.SetProcessDPIAware()
        except:
            pass
        
        dpg.setup_dearpygui()
        dpg.show_viewport()
        dpg.set_primary_window("Primary Window", True)

        # Start main loop
        while dpg.is_dearpygui_running():
            # Update progress bar if sending
            if self.is_sending:
                dpg.set_value("SendProgress", self.send_progress)
            
            dpg.render_dearpygui_frame()
        
        # Save settings on exit
        self.save_settings()
        dpg.destroy_context()


# Run the application
if __name__ == "__main__":
    print("Starting Sales Email Automator...")
    app = EmailAutomationApp()
    app.run()
