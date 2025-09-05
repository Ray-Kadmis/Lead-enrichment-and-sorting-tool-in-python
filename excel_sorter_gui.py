import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import pandas as pd
import re
import os
import threading
import time
import phonenumbers
import requests
from bs4 import BeautifulSoup
from urllib.parse import urlparse, urljoin
from collections import defaultdict
import random
import urllib3

# Disable SSL warnings
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

class ExcelSorterGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel/CSV Sorter Tool")
        self.root.geometry("800x600")
        self.root.resizable(True, True)
        
        # Variables
        self.selected_files = []
        self.processing = False
        
        # Create GUI
        self.create_widgets()
        
        # Center window
        self.center_window()
    
    def center_window(self):
        """Center the window on screen"""
        self.root.update_idletasks()
        x = (self.root.winfo_screenwidth() // 2) - (self.root.winfo_width() // 2)
        y = (self.root.winfo_screenheight() // 2) - (self.root.winfo_height() // 2)
        self.root.geometry(f"+{x}+{y}")
    
    def create_widgets(self):
        """Create all GUI widgets"""
        # Main frame
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configure grid weights
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(2, weight=1)
        main_frame.rowconfigure(4, weight=1)
        
        # Title
        title_label = ttk.Label(main_frame, text="Excel/CSV Sorter Tool", 
                               font=('Arial', 16, 'bold'))
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 20))
        
        # File selection section
        file_frame = ttk.LabelFrame(main_frame, text="File Selection", padding="10")
        file_frame.grid(row=1, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        file_frame.columnconfigure(0, weight=1)
        
        # Select files button
        select_btn = ttk.Button(file_frame, text="Select Excel/CSV Files", 
                               command=self.select_files, width=20)
        select_btn.grid(row=0, column=0, padx=(0, 10))
        
        # Clear files button
        clear_btn = ttk.Button(file_frame, text="Clear Selection", 
                              command=self.clear_files, width=15)
        clear_btn.grid(row=0, column=1)
        
        # Selected files listbox
        files_label = ttk.Label(file_frame, text="Selected Files:")
        files_label.grid(row=1, column=0, columnspan=2, sticky=tk.W, pady=(10, 5))
        
        # Listbox with scrollbar
        listbox_frame = ttk.Frame(file_frame)
        listbox_frame.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        listbox_frame.columnconfigure(0, weight=1)
        listbox_frame.rowconfigure(0, weight=1)
        
        self.files_listbox = tk.Listbox(listbox_frame, height=6)
        self.files_listbox.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        files_scrollbar = ttk.Scrollbar(listbox_frame, orient=tk.VERTICAL, 
                                       command=self.files_listbox.yview)
        files_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        self.files_listbox.configure(yscrollcommand=files_scrollbar.set)
        
        # Processing options
        options_frame = ttk.LabelFrame(main_frame, text="Processing Options", padding="10")
        options_frame.grid(row=2, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        
        # Combine files checkbox
        self.combine_var = tk.BooleanVar()
        combine_check = ttk.Checkbutton(options_frame, text="Combine all files into one", 
                                       variable=self.combine_var,
                                       command=self.toggle_combine_options)
        combine_check.grid(row=0, column=0, sticky=tk.W)
        
        # Output filename entry (for combine mode)
        self.output_frame = ttk.Frame(options_frame)
        self.output_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(10, 0))
        self.output_frame.columnconfigure(1, weight=1)
        
        ttk.Label(self.output_frame, text="Output filename:").grid(row=0, column=0, padx=(0, 10))
        self.output_entry = ttk.Entry(self.output_frame)
        self.output_entry.grid(row=0, column=1, sticky=(tk.W, tk.E))
        self.output_entry.insert(0, "Combined_Cleaned.xlsx")
        
        # Initially hide output options
        self.output_frame.grid_remove()
        
        # Process and Fetch Info buttons
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=3, column=0, columnspan=3, pady=20)
        
        # Process button
        self.process_btn = ttk.Button(button_frame, text="Process Files", 
                                    command=self.process_files, 
                                    style='Accent.TButton',
                                    width=15)
        self.process_btn.pack(side=tk.LEFT, padx=5)
        
        # Fetch Info button
        self.fetch_btn = ttk.Button(button_frame, text="Fetch Website Info",
                                   command=self.fetch_website_info,
                                   style='Accent.TButton',
                                   width=15)
        self.fetch_btn.pack(side=tk.LEFT, padx=5)
        
        # Progress bar
        self.progress = ttk.Progressbar(main_frame, mode='indeterminate')
        self.progress.grid(row=4, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        
        # Log/Output area
        log_frame = ttk.LabelFrame(main_frame, text="Processing Log", padding="10")
        log_frame.grid(row=5, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)
        
        self.log_text = scrolledtext.ScrolledText(log_frame, height=10, width=70)
        self.log_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Status bar
        self.status_var = tk.StringVar()
        self.status_var.set("Ready")
        status_bar = ttk.Label(main_frame, textvariable=self.status_var, 
                              relief=tk.SUNKEN, anchor=tk.W)
        status_bar.grid(row=6, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(10, 0))
    
    def select_files(self):
        """Open file dialog to select Excel/CSV files"""
        filetypes = [
            ('Excel and CSV files', '*.xlsx *.xls *.csv'),
            ('Excel files', '*.xlsx *.xls'),
            ('CSV files', '*.csv'),
            ('All files', '*.*')
        ]
        
        files = filedialog.askopenfilenames(
            title="Select Excel/CSV Files",
            filetypes=filetypes
        )
        
        if files:
            self.selected_files = list(files)
            self.update_files_listbox()
            self.log(f"Selected {len(files)} file(s)")
            self.status_var.set(f"{len(files)} file(s) selected")
    
    def clear_files(self):
        """Clear selected files"""
        self.selected_files = []
        self.update_files_listbox()
        self.log("File selection cleared")
        self.status_var.set("Ready")
    
    def update_files_listbox(self):
        """Update the files listbox with selected files"""
        self.files_listbox.delete(0, tk.END)
        for file in self.selected_files:
            filename = os.path.basename(file)
            self.files_listbox.insert(tk.END, filename)
    
    def toggle_combine_options(self):
        """Show/hide combine options based on checkbox"""
        if self.combine_var.get():
            self.output_frame.grid()
        else:
            self.output_frame.grid_remove()
    
    def log(self, message):
        """Add message to log"""
        self.log_text.insert(tk.END, f"{message}\n")
        self.log_text.see(tk.END)
        self.root.update_idletasks()
    
    def process_files(self):
        """Process selected files"""
        if not self.selected_files:
            messagebox.showwarning("No Files", "Please select files to process first.")
            return
        
        if self.processing:
            return
        
        # Start processing in separate thread
        self.processing = True
        self.process_btn.configure(state='disabled', text='Processing...')
        self.fetch_btn.configure(state='disabled')
        self.progress.start()
        
        thread = threading.Thread(target=self._process_files_thread)
        thread.daemon = True
        thread.start()
    
    def _process_files_thread(self):
        """Thread function for processing files"""
        try:
            if not self.selected_files:
                self.log("No files selected for processing")
                return
                
            sorter = ExcelSorter(log_callback=self.log)
            
            if len(self.selected_files) == 1 or not self.combine_var.get():
                # Process files individually
                success_count = 0
                for file_path in self.selected_files:
                    if sorter.process_single_file(file_path):
                        success_count += 1
                
                self.log(f"\nProcessing complete. Successfully processed {success_count} of {len(self.selected_files)} files.")
                
            else:
                # Combine files
                output_file = self.output_entry.get().strip()
                if not output_file:
                    output_file = "Combined_Cleaned.xlsx"
                elif not (output_file.endswith('.xlsx') or output_file.endswith('.xls') or output_file.endswith('.csv')):
                    output_file += ".xlsx"
                
                output_path = os.path.join(os.path.dirname(self.selected_files[0]), output_file)
                
                if sorter.process_multiple_files(self.selected_files, output_path):
                    self.log(f"\nSuccessfully combined and processed {len(self.selected_files)} files into: {output_path}")
                else:
                    self.log("\nError combining files. Please check the log for details.")
            
        except Exception as e:
            self.log(f"Error in processing thread: {str(e)}")
        finally:
            self.processing = False
            self.process_btn.configure(state='normal', text='Process Files')
            self.fetch_btn.configure(state='normal')
            self.progress.stop()
            self.status_var.set("Ready")
    
    def _fetch_website_info_thread(self, file_path):
        """Thread function for fetching website information"""
        try:
            if not file_path:
                self.log("No file selected for fetching website info")
                return
                
            self.log(f"Starting to fetch website information from: {file_path}")
            
            # Load the file
            sorter = ExcelSorter(log_callback=self.log)
            df = sorter.load_file(file_path)
            if df is None:
                self.log("Error: Could not load the file")
                return
                
            # Check if website column exists
            website_columns = [col for col in df.columns if 'website' in col.lower() or 'url' in col.lower()]
            if not website_columns:
                self.log("Error: No column containing 'website' or 'URL' found in the file")
                return
                
            website_column = website_columns[0]  # Use the first matching column
            self.log(f"Using column '{website_column}' for website URLs")
            
            # Fetch website information
            result_df = sorter.fetch_website_info_for_df(df, website_column)
            
            # Save the result
            base, ext = os.path.splitext(file_path)
            output_path = f"{base}_With_Contact_Info{ext}"
            
            if file_path.endswith('.csv'):
                result_df.to_csv(output_path, index=False)
            else:
                result_df.to_excel(output_path, index=False, engine='openpyxl')
                
            self.log(f"\nSuccessfully saved results to: {output_path}")
            self.log("\nSummary of added information:")
            self.log(f"- Email addresses: {len(result_df[result_df['Email_Addresses'] != ''])} rows")
            self.log(f"- Phone numbers: {len(result_df[result_df['Phone_Numbers'] != ''])} rows")
            
            # Count social media links
            social_cols = [col for col in result_df.columns if '_URL' in col and col != 'website_URL']
            for col in social_cols:
                count = len(result_df[result_df[col] != ''])
                if count > 0:
                    self.log(f"- {col.replace('_URL', '')} links: {count}")
            
        except Exception as e:
            self.log(f"Error in fetch website info thread: {str(e)}")
        finally:
            self.processing = False
            self.process_btn.configure(state='normal')
            self.fetch_btn.configure(state='normal', text='Fetch Website Info')
            self.progress.stop()
            self.status_var.set("Ready")
    
    def fetch_website_info(self):
        """Handle the Fetch Website Info button click"""
        if not self.selected_files:
            messagebox.showwarning("No Files", "Please select a file first.")
            return
            
        if self.processing:
            return
            
        if len(self.selected_files) > 1:
            messagebox.showinfo("Info", "Please select only one file at a time for fetching website information.")
            return
            
        # Confirm before proceeding
        if not messagebox.askyesno("Confirm", "This will fetch contact information from the websites in your file. "
                                           "This may take a while. Continue?"):
            return
        
        # Start processing in separate thread
        self.processing = True
        self.process_btn.configure(state='disabled')
        self.fetch_btn.configure(state='disabled', text='Fetching...')
        self.progress.start()
        self.status_var.set("Fetching website information...")
        
        thread = threading.Thread(
            target=self._fetch_website_info_thread,
            args=(self.selected_files[0],)
        )
        thread.daemon = True
        thread.start()

class ExcelSorter:
    def __init__(self, log_callback=None):
        self.required_columns = ['reviews', 'website', 'rating']
        self.log = log_callback if log_callback else print
        # Static pool of common desktop browser User-Agent strings to avoid fake-useragent dependency
        self.USER_AGENTS = [
            # Chrome (Windows)
            'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36',
            # Edge (Windows)
            'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36 Edg/124.0.0.0',
            # Firefox (Windows)
            'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:125.0) Gecko/20100101 Firefox/125.0',
            # Chrome (Mac)
            'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36',
            # Safari (Mac)
            'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/17.4 Safari/605.1.15',
        ]
        self.session = requests.Session()
        self.session.verify = False  # Disable SSL verification
        self.session.headers.update({
            'User-Agent': random.choice(self.USER_AGENTS),
            'Accept-Language': 'en-US,en;q=0.5',
            'Accept-Encoding': 'gzip, deflate, br',
            'Connection': 'keep-alive',
            'Upgrade-Insecure-Requests': '1',
            'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8'
        })
    
    def _is_valid_url(self, url):
        """Check if the URL is valid"""
        if not url or pd.isna(url) or not isinstance(url, str):
            return False
        return url.startswith(('http://', 'https://'))
        
    def extract_domain(self, url):
        """Extract domain name from URL"""
        if not self._is_valid_url(url):
            return None
        
        try:
            # Clean the URL
            url = str(url).strip()
            
            # Add protocol if missing
            if not url.startswith(('http://', 'https://')):
                url = 'http://' + url
            
            # Parse URL
            parsed = urlparse(url)
            domain = parsed.netloc.lower()
            
            # Remove www. prefix
            if domain.startswith('www.'):
                domain = domain[4:]
            
            # Extract main domain name (before first dot)
            domain_parts = domain.split('.')
            if len(domain_parts) >= 2:
                return domain_parts[0]
            return domain
        except Exception as e:
            self.log(f"Error extracting domain from {url}: {str(e)}")
            return None
    
    def _get_page_content(self, url, timeout=10, max_retries=2):
        """Get page content with retries"""
        for attempt in range(max_retries):
            try:
                # Rotate a realistic User-Agent for each request
                self.session.headers['User-Agent'] = random.choice(self.USER_AGENTS)
                response = self.session.get(url, timeout=timeout, allow_redirects=True)
                response.raise_for_status()
                return response.text
            except requests.RequestException as e:
                if attempt == max_retries - 1:
                    self.log(f"Failed to fetch {url}: {str(e)}")
                    return None
                time.sleep(1)  # Wait before retry
    
    def _extract_emails(self, text):
        """Extract email addresses from text"""
        if not text:
            return set()
        # Improved email pattern to find more variations
        email_pattern = r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b'
        emails = set()
        for match in re.finditer(email_pattern, text, re.IGNORECASE):
            email = match.group(0).strip()
            if '.' in email.split('@')[-1]:  # Must have a dot in the domain part
                emails.add(email)
        return emails
    
    def _extract_phone_numbers(self, text, default_region='US'):
        """Extract and validate phone numbers from text"""
        if not text:
            return set()
            
        # First, try to find phone numbers using common patterns
        phone_patterns = [
            r'\+?\d{1,4}?[-.\s]?\(?\d{1,4}?\)?[-.\s]?\d{1,4}[-.\s]?\d{1,4}[-.\s]?\d{1,9}',  # International
            r'\(?\d{3}\)?[-.\s]?\d{3}[-.\s]?\d{4}',  # US/Canada
            r'\d{3}[-.\s]?\d{3}[-.\s]?\d{4}'  # US/Canada without area code
        ]
        
        found_numbers = set()
        for pattern in phone_patterns:
            for match in re.finditer(pattern, text):
                try:
                    phone = match.group(0)
                    # Clean up the phone number
                    phone = re.sub(r'[^\d+]', '', phone)
                    if phone.startswith('00'):
                        phone = '+' + phone[2:]
                    elif phone.startswith('1') and len(phone) == 11 and not phone.startswith('+1'):
                        phone = '+1' + phone[1:]
                    elif not phone.startswith('+'):
                        phone = '+1' + phone  # Default to US/Canada
                        
                    # Parse and validate the phone number
                    parsed = phonenumbers.parse(phone, None)
                    if phonenumbers.is_valid_number(parsed):
                        formatted = phonenumbers.format_number(
                            parsed, 
                            phonenumbers.PhoneNumberFormat.INTERNATIONAL
                        )
                        found_numbers.add(formatted)
                except Exception:
                    continue
        
        return found_numbers
    
    def _extract_social_links(self, soup, base_url):
        """Extract social media links from the page"""
        social_platforms = {
            'facebook.com': 'Facebook',
            'fb.com': 'Facebook',
            'twitter.com': 'Twitter',
            'x.com': 'Twitter',  # Twitter's new domain
            'linkedin.com': 'LinkedIn',
            'instagram.com': 'Instagram',
            'youtube.com': 'YouTube',
            'pinterest.com': 'Pinterest'
        }
        
        social_links = {}
        
        # Find all links on the page
        for a in soup.find_all('a', href=True):
            if not a.get_text(strip=True):
                continue  # Skip empty links
                
            href = a['href'].lower()
            
            # Check for direct social media profile links
            for domain, platform in social_platforms.items():
                if domain in href:
                    full_url = urljoin(base_url, href)
                    # Clean up the URL
                    full_url = full_url.split('?')[0]  # Remove query parameters
                    full_url = full_url.rstrip('/')  # Remove trailing slash
                    social_links[platform] = full_url
                    break
            
            # Check for social media icons/buttons
            if 'social' in a.get('class', []) or 'social' in a.get('id', '').lower():
                for domain, platform in social_platforms.items():
                    if domain in href:
                        full_url = urljoin(base_url, href)
                        social_links[platform] = full_url
                        break
        
        return social_links
    
    def _find_contact_page_links(self, soup, base_url):
        """Find links to contact, about, or info pages"""
        contact_links = set()
        
        # Common contact page patterns
        contact_keywords = [
            'contact', 'about', 'info', 'reach', 'connect', 'get in touch',
            'contact us', 'about us', 'get in contact', 'find us', 'reach us'
        ]
        
        # Check all links on the page
        for a in soup.find_all('a', href=True):
            href = a.get('href', '').lower()
            text = a.get_text(' ', strip=True).lower()
            
            # Check if link text or URL contains contact keywords
            if any(keyword in text or keyword in href for keyword in contact_keywords):
                full_url = urljoin(base_url, href)
                contact_links.add(full_url)
        
        return list(contact_links)[:3]  # Return first 3 unique contact links
    
    def _extract_facebook_info(self, soup, base_url):
        """Extract information from Facebook pages"""
        emails = set()
        
        try:
            # Try to find the intro section
            intro_section = soup.find('div', {'id': 'intro_container_id'}) or \
                          soup.find('div', class_=re.compile(r'(?i)intro|about|bio|description'))
            
            if intro_section:
                intro_text = intro_section.get_text(' ')
                emails.update(self._extract_emails(intro_text))
            
            # Look for the 'About' section
            about_links = [a['href'] for a in soup.find_all('a', href=True) 
                         if 'about' in a.get('href', '').lower() 
                         and 'profile.php' not in a.get('href', '')]
            
            if about_links:
                about_url = urljoin(base_url, about_links[0])
                about_content = self._get_page_content(about_url)
                
                if about_content:
                    about_soup = BeautifulSoup(about_content, 'lxml')
                    
                    # Look for contact information sections
                    contact_sections = about_soup.find_all(['div', 'section'], 
                                                         class_=re.compile(r'(?i)contact|info|details'))
                    
                    for section in contact_sections:
                        section_text = section.get_text(' ')
                        emails.update(self._extract_emails(section_text))
            
            return emails
            
        except Exception as e:
            self.log(f"Error extracting Facebook info: {str(e)}")
            return emails
    
    def scrape_website_info(self, url):
        """Scrape contact information from a website"""
        if not self._is_valid_url(url):
            return {'error': 'Invalid URL'}
        
        self.log(f"Scraping: {url}")
        
        try:
            # Get the main page content
            content = self._get_page_content(url)
            if not content:
                return {'error': 'Could not fetch page content'}
                
            soup = BeautifulSoup(content, 'lxml')
            
            # Extract emails and phone numbers from the main page
            text = soup.get_text(' ')
            emails = self._extract_emails(text)
            phones = self._extract_phone_numbers(text)
            
            # Check if this is a social media profile
            is_facebook = any(domain in url.lower() for domain in ['facebook.com', 'fb.com'])
            is_instagram = 'instagram.com' in url.lower()
            is_linkedin = 'linkedin.com' in url.lower()
            
            # Special handling for social media profiles
            if is_facebook:
                facebook_emails = self._extract_facebook_info(soup, url)
                emails.update(facebook_emails)
            elif is_instagram or is_linkedin:
                # For Instagram and LinkedIn, look for bio/description
                bio_section = soup.find('div', class_=re.compile(r'(?i)bio|description|about'))
                if bio_section:
                    bio_text = bio_section.get_text(' ')
                    emails.update(self._extract_emails(bio_text))
            else:
                # For regular websites, look for contact pages
                contact_links = self._find_contact_page_links(soup, url)
                
                # Check contact pages for more information
                for contact_link in contact_links:
                    try:
                        contact_content = self._get_page_content(contact_link)
                        if contact_content:
                            contact_soup = BeautifulSoup(contact_content, 'lxml')
                            contact_text = contact_soup.get_text(' ')
                            
                            # Extract emails and phones from contact page
                            emails.update(self._extract_emails(contact_text))
                            phones.update(self._extract_phone_numbers(contact_text))
                            
                            # Look for email links
                            for a in contact_soup.find_all('a', href=True):
                                if 'mailto:' in a['href'].lower():
                                    email = a['href'].replace('mailto:', '').strip()
                                    if '@' in email and '.' in email:
                                        emails.add(email)
                    except Exception as e:
                        self.log(f"Error checking contact page {contact_link}: {str(e)}")
            
            # Extract social media links
            social_links = self._extract_social_links(soup, url)
            
            # Format the results
            result = {
                'emails': sorted(list(emails)),
                'phone_numbers': sorted(list(phones)),
                'social_links': social_links,
                'status': 'Success',
                'website': url  # Include the website URL in the result
            }
            
            return result
            
        except Exception as e:
            self.log(f"Error scraping {url}: {str(e)}")
            return {'error': str(e), 'website': url}

    
    def find_columns(self, df):
        """Find required columns in dataframe (case insensitive)"""
        column_mapping = {}
        df_columns_lower = [col.lower() for col in df.columns]
        
        for required_col in self.required_columns:
            found = False
            for i, col in enumerate(df_columns_lower):
                if required_col in col or col in required_col:
                    column_mapping[required_col] = df.columns[i]
                    found = True
                    break
            
            if not found:
                self.log(f"Warning: Column '{required_col}' not found in the data")
                return None
        
        return column_mapping
    
    def process_dataframe(self, df):
        """Process the dataframe according to requirements"""
        # Find required columns
        column_mapping = self.find_columns(df)
        if not column_mapping:
            return None
        
        # Create working copy
        df_work = df.copy()
        
        # Rename columns for easier processing
        website_col = column_mapping['website']
        reviews_col = column_mapping['reviews']
        rating_col = column_mapping['rating']
        
        # Convert reviews to numeric, handling errors
        df_work[reviews_col] = pd.to_numeric(df_work[reviews_col], errors='coerce').fillna(0)
        
        # Step 1: Separate rows with empty websites
        empty_website_mask = df_work[website_col].isna() | (df_work[website_col] == '') | (df_work[website_col] == 'nan')
        empty_website_rows = df_work[empty_website_mask].copy()
        non_empty_website_rows = df_work[~empty_website_mask].copy()
        
        # Step 2: Sort empty website rows by reviews (highest first)
        empty_website_rows = empty_website_rows.sort_values(by=reviews_col, ascending=False)
        
        # Step 3: Process non-empty website rows for domain extraction
        non_empty_website_rows['domain'] = non_empty_website_rows[website_col].apply(self.extract_domain)
        
        # Group by domain to find repeating businesses
        domain_groups = defaultdict(list)
        single_domain_rows = []
        
        for idx, row in non_empty_website_rows.iterrows():
            domain = row['domain']
            if domain:
                domain_groups[domain].append(idx)
            else:
                single_domain_rows.append(idx)
        
        # Separate repeated and single domain businesses
        repeated_businesses = []
        single_businesses = []
        
        for domain, indices in domain_groups.items():
            if len(indices) > 1:
                # Sort repeated businesses by reviews within each domain group
                domain_rows = non_empty_website_rows.loc[indices].sort_values(by=reviews_col, ascending=False)
                repeated_businesses.append(domain_rows)
            else:
                single_businesses.extend(indices)
        
        # Add single domain rows
        if single_domain_rows:
            single_businesses.extend(single_domain_rows)
        
        # Get single business rows and sort by reviews
        single_business_rows = non_empty_website_rows.loc[single_businesses].sort_values(by=reviews_col, ascending=False)
        
        # Step 4: Combine all data
        result_df = pd.concat([empty_website_rows, single_business_rows], ignore_index=True)
        
        # Add repeated businesses section
        if repeated_businesses:
            # Add separator row
            separator_row = pd.DataFrame([[''] * len(df.columns)], columns=df.columns)
            separator_row.iloc[0, 0] = 'Repeated Businesses'
            
            result_df = pd.concat([result_df, separator_row], ignore_index=True)
            
            # Add repeated business groups
            for domain_group in repeated_businesses:
                domain_group_clean = domain_group.drop('domain', axis=1)
                result_df = pd.concat([result_df, domain_group_clean], ignore_index=True)
        
        # Remove the temporary domain column if it exists
        if 'domain' in result_df.columns:
            result_df = result_df.drop('domain', axis=1)
        
        return result_df
    
    def load_file(self, file_path):
        """Load Excel or CSV file"""
        try:
            if file_path.lower().endswith('.csv'):
                return pd.read_csv(file_path)
            else:
                return pd.read_excel(file_path)
        except Exception as e:
            self.log(f"Error loading file {file_path}: {str(e)}")
            return None
    
    def process_single_file(self, input_file, output_dir=None):
        """Process a single file"""
        try:
            self.log(f"Processing file: {input_file}")
            df = self.load_file(input_file)
            if df is None:
                return False
                
            # Process the dataframe
            processed_df = self.process_dataframe(df)
            
            # Save the result
            if output_dir:
                os.makedirs(output_dir, exist_ok=True)
                output_path = os.path.join(output_dir, f"Cleaned_{os.path.basename(input_file)}")
            else:
                base, ext = os.path.splitext(input_file)
                output_path = f"{base}_Cleaned{ext}"
                
            if input_file.endswith('.csv'):
                processed_df.to_csv(output_path, index=False)
            else:
                processed_df.to_excel(output_path, index=False, engine='openpyxl')
                
            self.log(f"Saved cleaned file to: {output_path}")
            return True
            
        except Exception as e:
            self.log(f"Error processing file {input_file}: {str(e)}")
            return False
            
    def fetch_website_info_for_df(self, df, website_column='website'):
        """Fetch website information for all websites in the dataframe"""
        if website_column not in df.columns:
            self.log(f"Error: Column '{website_column}' not found in the dataframe")
            return df
            
        # Add new columns if they don't exist
        new_columns = {
            'Email_Addresses': [],
            'Phone_Numbers': [],
            'Facebook_URL': [],
            'Instagram_URL': [],
            'LinkedIn_URL': [],
            'Twitter_URL': [],
            'YouTube_URL': [],
            'Pinterest_URL': []
        }
        
        # Initialize new columns with empty values
        for col in new_columns.keys():
            if col not in df.columns:
                df[col] = ''
        
        # Process each row
        total_rows = len(df)
        for idx, row in df.iterrows():
            url = row[website_column]
            if not self._is_valid_url(url):
                self.log(f"Skipping invalid URL at row {idx + 2}: {url}")
                for col in new_columns.keys():
                    new_columns[col].append('')
                continue
                
            # Scrape website info
            result = self.scrape_website_info(url)
            
            # Update the row with scraped data
            if 'error' in result:
                self.log(f"Error processing {url}: {result['error']}")
                for col in new_columns.keys():
                    new_columns[col].append('')
                continue
                
            # Update emails
            emails = result.get('emails', [])
            df.at[idx, 'Email_Addresses'] = ', '.join(emails) if emails else ''
            
            # Update phone numbers
            phones = result.get('phone_numbers', [])
            df.at[idx, 'Phone_Numbers'] = ' | '.join(phones) if phones else ''
            
            # Update social media links
            social_links = result.get('social_links', {})
            for platform, url in social_links.items():
                col_name = f"{platform}_URL"
                if col_name in df.columns:
                    df.at[idx, col_name] = url
            
            # Log progress
            if (idx + 1) % 5 == 0 or (idx + 1) == total_rows:
                self.log(f"Processed {idx + 1}/{total_rows} rows")
            
            # Be nice to servers
            time.sleep(1)
        
        return df
    
    def process_multiple_files(self, input_files, output_file="Combined_Cleaned.xlsx"):
        """Process multiple files and combine into one"""
        self.log(f"Processing {len(input_files)} files...")
        
        all_data = []
        
        for file_path in input_files:
            self.log(f"Loading: {os.path.basename(file_path)}")
            df = self.load_file(file_path)
            if df is not None:
                processed_df = self.process_dataframe(df)
                if processed_df is not None:
                    all_data.append(processed_df)
                else:
                    self.log(f"Skipping {file_path} - missing required columns")
            else:
                self.log(f"Skipping {file_path} - could not load")
        
        if not all_data:
            self.log("No valid files to process")
            return False
        
        # Combine all dataframes
        combined_df = pd.concat(all_data, ignore_index=True)
        
        # Process the combined data again to handle cross-file duplicates
        final_df = self.process_dataframe(combined_df)
        
        # Save combined file
        try:
            final_df.to_excel(output_file, index=False)
            self.log(f"✓ Combined file saved: {output_file}")
            return True
        except Exception as e:
            self.log(f"Error saving combined file: {str(e)}")
            return False

def main():
    root = tk.Tk()
    app = ExcelSorterGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()
