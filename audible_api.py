import os
import json
import sys
import asyncio
import requests
from getpass import getpass
import webbrowser
import io
from datetime import datetime

import pandas as pd
import pandas.io.formats.excel
import audible
import httpx

from pydub import AudioSegment

import speech_recognition as sr

from errors import ExternalError
from constants import artifacts_root_directory

# not currently in use, but so the user can choose their store
country_code_mapping = {
    "us": ".com",
    "ca": ".ca",
    "uk": ".co.uk",
    "au": ".com.au",
    "fr": ".fr",
    "de": ".de",
    "jp": ".co.jp",
    "it": ".it",
    "in": ".co.in",
    "es": ".es"
}

AUDIBLE_URL_BASE = "https://www.audible"

# set in ms, how long before and after the bookmark timestamp we want to slice the audioclips, useful for redundancy
# i.e to account for the time the user spends to dig up their phone and click bookmark
# Feel free to vary these, but free Speech Recognition API's have certain limits...
START_POSITION_OFFSET = 10000
END_POSITION_OFFSET = 0

class AudibleAPI:

    def __init__(self, auth):
        self.auth = auth
        self.books = []
        self.library = {}

    @classmethod
    async def authenticate(cls) -> "AudibleAPI":
        secrets_dir_path = os.path.join(artifacts_root_directory, "secrets")
        credentials_path = os.path.join(secrets_dir_path, "credentials.json")
        
        if os.path.exists(credentials_path):
            print(f"You are already authenticated, to switch accounts, delete secrets directory under {artifacts_root_directory} and try again")
            return None
            
        print("=== Audible Authentication ===")
        print("This will authenticate with your Audible account.")
        print("Note: Two-Factor Authentication (2FA) must be enabled on your Amazon account.")
        print()
        
        return await cls._authenticate_with_browser_assistance(secrets_dir_path, credentials_path)

    @classmethod
    async def _authenticate_with_browser_assistance(cls, secrets_dir_path, credentials_path):
        """Enhanced authentication with browser assistance for 2FA issues"""
        print("=== Enhanced Authentication Process ===")
        print("If you're having trouble with 2FA codes, try this process:")
        print("1. First, log into Amazon.com in your browser to verify your account is working")
        print("2. Make sure 2FA is properly configured in your Amazon account settings")
        print("3. Have your phone/email ready for receiving verification codes")
        print()
        
        browser_prep = input("Would you like me to open Amazon.com in your browser first? [y/N]: ").strip().lower()
        if browser_prep == 'y':
            print("Opening Amazon.com in your browser...")
            try:
                webbrowser.open("https://amazon.com")
                print("Please log into Amazon.com in your browser to verify your account is working.")
                input("Press Enter when you've successfully logged into Amazon.com...")
            except Exception as e:
                print(f"Could not open browser: {e}")
                print("Please manually visit https://amazon.com and log in, then come back here.")
                input("Press Enter when ready to continue...")
        
        email = input("\nAudible Email: ")
        password = getpass("Enter Password (will be hidden, press ENTER when done): ")
        
        print("\nAvailable regions:")
        print(', '.join(country_code_mapping.keys()))
        locale = input("Please enter your locale from the list above: ")
        
        if locale not in country_code_mapping:
            print(f"Invalid locale. Please choose from: {', '.join(country_code_mapping.keys())}")
            return None

        def captcha_callback(captcha_url: str) -> str:
            """Helper function for handling captcha with browser support."""
            print(f"\n=== CAPTCHA Required ===")
            print(f"A CAPTCHA challenge has been presented.")
            
            try:
                # Try to open the CAPTCHA in the default browser
                print("Opening CAPTCHA in your default browser...")
                webbrowser.open(captcha_url)
                print("Please view the CAPTCHA in your browser and enter the solution below.")
            except Exception as e:
                print(f"Could not open browser automatically: {e}")
                print(f"Please manually open this URL in your browser:")
                print(f"{captcha_url}")
            
            while True:
                guess = input("Enter CAPTCHA solution: ").strip()
                if guess:
                    return guess
                print("Please enter a valid CAPTCHA solution.")

        def cvf_callback() -> str:
            """Helper function for handling CVF (Challenge Verification Form) codes."""
            print(f"\n=== Two-Factor Authentication Required ===")
            print(f"🔍 DEBUG: CVF callback has been triggered by the Audible library")
            print(f"🔍 DEBUG: This means Amazon is requesting 2FA verification")
            print(f"🔍 DEBUG: Timestamp: {datetime.now().strftime('%H:%M:%S')}")
            print("Amazon is requesting a verification code for two-factor authentication.")
            print()
            
            # Enhanced troubleshooting for 2FA issues
            print("🔧 TROUBLESHOOTING: Not receiving codes?")
            print("1. Check your Amazon account 2FA settings at: https://amazon.com/myaccount")
            print("2. Ensure your phone number and email are correct and verified")
            print("3. Try requesting a new code if the current one expires")
            print("4. Check spam/junk folders for email codes")
            print("5. Make sure your phone has good signal for SMS codes")
            print()
            
            print("You should receive a verification code via one of these methods:")
            print("• SMS text message to your registered phone number")
            print("• Email to your registered email address") 
            print("• Push notification through the Amazon mobile app")
            print("• Authentication app (like Google Authenticator)")
            print()
            
            # Offer to open Amazon account settings
            settings_help = input("Open Amazon account settings in browser to check 2FA setup? [y/N]: ").strip().lower()
            if settings_help == 'y':
                try:
                    amazon_url = f"https://amazon{country_code_mapping.get(locale, '.com')}/myaccount"
                    print(f"🔍 DEBUG: Opening URL: {amazon_url}")
                    webbrowser.open(amazon_url)
                    print(f"Opened {amazon_url} in your browser.")
                    print("Check your 2FA settings and make sure they're properly configured.")
                except Exception as e:
                    print(f"Could not open browser: {e}")
                    print(f"Please manually visit: https://amazon{country_code_mapping.get(locale, '.com')}/myaccount")
            
            print("\n🔍 DEBUG: Amazon should be sending you a verification code now...")
            print("🔍 DEBUG: This usually happens immediately after you enter your password")
            print("🔍 DEBUG: Check ALL possible delivery methods (SMS, email, app notifications)")
            print()
            print("💡 TIP: If you don't receive a code within 1-2 minutes:")
            print("   1. Try logging into Amazon.com in a browser with the same credentials")
            print("   2. This will show you what 2FA method Amazon is actually using")
            print("   3. You might need to update your 2FA settings in Amazon")
            print()
            
            # Add option to check Amazon login in browser
            browser_check = input("Want me to open Amazon login in browser to test your 2FA setup? [y/N]: ").strip().lower()
            if browser_check == 'y':
                try:
                    login_url = f"https://amazon{country_code_mapping.get(locale, '.com')}/ap/signin"
                    print(f"🔍 DEBUG: Opening Amazon login: {login_url}")
                    webbrowser.open(login_url)
                    print("Try logging in with the same email/password and see what 2FA options appear.")
                    input("Press Enter when you've tested the login in your browser...")
                except Exception as e:
                    print(f"Could not open browser: {e}")
            
            print("\nWaiting for your verification code...")
            print("Tip: Codes are usually 6 digits and expire quickly (within a few minutes)")
            
            attempts = 0
            max_attempts = 5  # Increased from 3 to 5
            
            while attempts < max_attempts:
                print(f"\n🔍 DEBUG: Attempt {attempts + 1} of {max_attempts}")
                print(f"🔍 DEBUG: Current time: {datetime.now().strftime('%H:%M:%S')}")
                
                # Add more detailed prompting
                if attempts == 0:
                    prompt_msg = f"Enter the 6-digit verification code"
                elif attempts == 1:
                    prompt_msg = f"Try again - enter the verification code (check SMS/email)"
                elif attempts == 2:
                    prompt_msg = f"Still waiting - enter the code (check all devices/apps)"
                else:
                    prompt_msg = f"Last chances - enter the verification code"
                
                cvf_code = input(f"{prompt_msg}: ").strip()
                
                print(f"🔍 DEBUG: You entered: '{cvf_code}' (length: {len(cvf_code)})")
                
                if cvf_code and len(cvf_code) >= 4:  # Accept 4+ digits
                    print(f"🔍 DEBUG: Code format looks valid, returning to Audible library...")
                    return cvf_code
                
                attempts += 1
                print(f"🔍 DEBUG: Code format invalid (too short or empty)")
                
                if attempts < max_attempts:
                    print("Invalid code format. Please try again.")
                    
                    # Offer more specific help based on attempt number
                    if attempts == 1:
                        print("💡 TIP: Check your text messages and email (including spam folder)")
                    elif attempts == 2:
                        print("💡 TIP: Try checking Amazon mobile app notifications")
                        print("💡 TIP: Or try generating a new code if you have an authenticator app")
                    elif attempts == 3:
                        print("💡 TIP: The code might have expired, try requesting a new one")
                    
                    retry = input("Request a new code from Amazon? [y/N]: ").strip().lower()
                    if retry == 'y':
                        print("🔍 DEBUG: User requested new code generation")
                        print("Please request a new verification code from Amazon and try again.")
                        print("🔍 DEBUG: Wait for the new code before entering it...")
                else:
                    print("\nMax attempts reached. Please check your 2FA setup and try again later.")
                    print("🔍 DEBUG: Exceeded maximum attempts, raising exception")
                    print("Make sure:")
                    print("• 2FA is enabled on your Amazon account")
                    print("• Your phone number and email are verified")
                    print("• You're checking the right device/email for codes")
                    print("• Your account doesn't have any security restrictions")
                    raise Exception("Too many failed verification attempts")

        try:
            print("\nAttempting to authenticate with Audible...")
            print("🔍 DEBUG: Starting authentication process...")
            print("This may take a moment and might require solving a CAPTCHA or entering a 2FA code...")
            
            # Add timeout and retry logic
            max_auth_attempts = 3
            for auth_attempt in range(max_auth_attempts):
                try:
                    print(f"🔍 DEBUG: Authentication attempt {auth_attempt + 1} of {max_auth_attempts}")
                    print(f"🔍 DEBUG: Calling audible.Authenticator.from_login() at {datetime.now().strftime('%H:%M:%S')}")
                    
                    auth = audible.Authenticator.from_login(
                        email,
                        password,
                        locale=locale,
                        with_username=False,
                        captcha_callback=captcha_callback,
                        cvf_callback=cvf_callback
                    )
                    
                    print(f"🔍 DEBUG: Authentication successful at {datetime.now().strftime('%H:%M:%S')}")
                    break
                    
                except Exception as auth_error:
                    error_msg = str(auth_error).lower()
                    print(f"🔍 DEBUG: Authentication attempt {auth_attempt + 1} failed: {auth_error}")
                    
                    if "timeout" in error_msg or "timed out" in error_msg:
                        if auth_attempt < max_auth_attempts - 1:
                            wait_time = (auth_attempt + 1) * 10  # 10, 20, 30 seconds
                            print(f"⏱️  Network timeout detected. Waiting {wait_time} seconds before retry...")
                            print("💡 TIP: This is usually a temporary server issue, not a problem with your credentials")
                            import time
                            time.sleep(wait_time)
                            continue
                        else:
                            print("❌ Multiple timeout errors - this appears to be a server/network issue")
                            print("🔧 TROUBLESHOOTING TIPS:")
                            print("• Try again in a few minutes - Audible servers may be busy")
                            print("• Check your internet connection stability")
                            print("• Try using a different network (mobile hotspot, etc.)")
                            print("• The authentication worked but the server response was slow")
                            return None
                    else:
                        # Re-raise non-timeout errors immediately
                        raise auth_error
            else:
                # This should not happen due to the break, but just in case
                print("❌ All authentication attempts failed")
                return None
            
            print("🔍 DEBUG: Saving credentials to file...")
            os.makedirs(secrets_dir_path, exist_ok=True)
            auth.to_file(credentials_path)
            print("✅ Authentication successful!")
            print("✅ Credentials saved locally")
            print(f"🔍 DEBUG: Process completed at {datetime.now().strftime('%H:%M:%S')}")
            return cls(auth)
            
        except Exception as e:
            print(f"🔍 DEBUG: Final exception caught: {e}")
            # Handle all authentication errors generically since the specific exception classes
            # may vary between audible library versions
            error_message = str(e).lower()
            
            if any(keyword in error_message for keyword in ['login', 'authentication', 'credentials', 'password']):
                print(f"❌ Login failed: {e}")
                print("\n🔧 TROUBLESHOOTING TIPS:")
                print("• Verify your email and password are correct")
                print("• Ensure 2FA is properly set up on your Amazon account") 
                print("• Try logging into Amazon.com in a browser first")
                print("• Check that you're using the correct regional store")
                print("• Wait a few minutes and try again (rate limiting)")
            elif any(keyword in error_message for keyword in ['verification', 'cvf', '2fa', 'code']):
                print(f"❌ 2FA verification failed: {e}")
                print("\n🔧 TROUBLESHOOTING TIPS:")
                print("• Make sure you entered the correct verification code")
                print("• Request a new code if the previous one expired")
                print("• Check your 2FA setup in Amazon account settings")
                print("• Try the verification code as soon as you receive it")
            else:
                print(f"❌ Authentication failed: {e}")
                print("\n🔧 TROUBLESHOOTING TIPS:")
                print("• Check your internet connection")
                print("• Try again in a few minutes")
                print("• Verify your Amazon account is in good standing")
                print("• Make sure 2FA is properly configured")
            
            return None

    # Gets information about a book
    async def get_book_infos(self, asin):
        async with audible.AsyncClient(self.auth) as client:
            try:                
                book = await client.get(
                    path=f"library/{asin}",
                    params={
                        "response_groups": (
                            "contributors, media, price, reviews, product_attrs, "
                            "product_extended_attrs, product_desc, product_plan_details, "
                            "product_plans, rating, sample, sku, series, ws4v, origin, "
                            "relationships, review_attrs, categories, badge_types, "
                            "category_ladders, claim_code_url, is_downloaded, pdf_url, "
                            "is_returnable, origin_asin, percent_complete, provided_review"
                        )
                    }
                )
                return book
            except Exception as e:
                print(e)

    # Helper function for displaying the users books and allowing them to select one based on the index number
    async def get_book_selection(self):

        if not self.library:
            await self.get_library()

        li_books = []
        # if not self.lib
        for index, book in enumerate(self.library["items"]):
            li_books.append(book["asin"])
            book_title = book.get("title", "Unable to retrieve book name")
            print(f"{index}: {book_title}")

        book_selection = input(
            "Enter the index number of the book you would like to download, or enter --all for all available books: \n")

        if book_selection == "--all":
            li_books = [{"title": book.get("title", 'untitled'), "asin": book["asin"]}
                        for book in self.library["items"]]

        else:
            try:
                li_books = [{"title": self.library["items"][int(book_selection)],
                             "asin":self.library["items"][int(book_selection)].get("asin", None)}]
            except (IndexError, ValueError):
                print("Invalid selection")                
        return li_books

    # Main download books function
    async def cmd_download_books(self):
        li_books = await self.get_book_selection()

        tasks = []
        for book in li_books:
            tasks.append(
                asyncio.ensure_future(
                    self.get_book_infos(
                        book.get("asin"))))

        books = await asyncio.gather(*tasks)

        all_books = {}

        for book in books:
            if book is not None:
                print(book["item"]["title"])
                asin = book["item"]["asin"]
                raw_title = book["item"]["title"]
                title = raw_title.lower().replace(" ", "_")
                all_books[asin] = title

                # Attempt to download book
                try:
                    re = self.get_download_url(self.generate_url(self.auth.locale.country_code, "download", asin), num_results=1000, response_groups="product_desc, product_attrs")

                # Audible API throws error, usually for free books that are not allowed to be downloaded, we skip to the next
                except audible.exceptions.NetworkError as e:
                    ExternalError(self.get_download_url,
                                  asin, e).show_error()
                    continue

                audible_response = requests.get(re, stream=True)

                title_dir_path = os.path.join(artifacts_root_directory, "audiobooks", title)
                path_exists = os.path.exists(title_dir_path)
                if not path_exists:
                    os.makedirs(title_dir_path)
                    

                if audible_response.ok:
                    title_file_path = os.path.join(title_dir_path, f"{title}.aax")
                    with open(title_file_path, 'wb') as f:
                        print("Downloading %s" % raw_title)

                        total_length = audible_response.headers.get(
                            'content-length')

                        if total_length is None:  # no content length header
                            print(
                                "Unable to estimate download size, downloading, this might take a while...")
                            f.write(audible_response.content)
                        else:
                            # Save book locally and calculate and print download progress (progress bar)
                            dl = 0
                            total_length = int(total_length)
                            for data in audible_response.iter_content(chunk_size=1024*1024):
                                dl += len(data)
                                f.write(data)
                                done = int(50 * dl / total_length)
                                sys.stdout.write("\r[%s%s]" % ('=' * done, ' ' * (50-done)))

                                sys.stdout.write(f"   {int(dl / total_length * 100)}%")
                                sys.stdout.flush()

                else:
                    print(audible_response.text)

    # WIP
    def generate_url(self, country_code, url_type, asin=None):
        if asin and url_type == "download":
            return f"{AUDIBLE_URL_BASE}{country_code_mapping.get(country_code)}/library/download?asin={asin}&codec=AAX"

    # Need the next_request for Audible API to give us the download link for the book
    def get_download_link_callback(self, resp):
        return resp.next_request

    # Sends a request to get the download link for the selected book
    def get_download_url(self, url, **kwargs):

        with audible.Client(auth=self.auth, response_callback=self.get_download_link_callback) as client:
            library = client.get(
                url,
                **kwargs
            )
            return library.url

    async def cmd_list_books(self):
        if not self.books:
            await self.cmd_show_library()

        await self.cmd_show_library()
        
    # Gets all books and info for account and adds it to self.books, also returns ASIN for all books
    async def get_library(self):
        async with audible.AsyncClient(self.auth) as client:
            self.library = await client.get(
                path="library",
                params={
                    "num_results": 999
                }
            )
            asins = [book["asin"] for book in self.library["items"]]

            for book in self.library["items"]:
                asins.append(book["asin"])
                book_title = book.get("title", "Unable to retrieve book name")
                self.books.append(book_title)

            return asins

    async def cmd_show_library(self):
        if not self.books:
            await self.get_library()

        for index, book_title in enumerate(self.books):
            print(f"{index}: {book_title}")
   

    async def cmd_get_bookmarks(self):
        li_books = await self.get_book_selection()

        for book in li_books:
            print(self.get_bookmarks(book))

    def get_bookmarks(self, book):
        asin = book.get("asin")
        # Handle both string and nested dictionary formats for title
        title_value = book.get("title", {})
        if isinstance(title_value, str):
            _title = title_value
        else:
            _title = title_value.get("title", "untitled")
        
        if not _title:
            return

        title = _title.lower().replace(" ", "_")

        bookmarks_url = f"https://cde-ta-g7g.amazon.com/FionaCDEServiceEngine/sidecar?type=AUDI&key={asin}"
        print(f"Getting bookmarks for {_title}")
        with audible.Client(auth=self.auth, response_callback=self.bookmark_response_callback) as client:
            library = client.get(
                bookmarks_url,
                num_results=1000,
                response_groups="product_desc, product_attrs"
            )

            li_bookmarks = library.json().get("payload", {}).get("records", [])
            li_clips = sorted(
                li_bookmarks, key=lambda i: i["type"], reverse=True)

            title_dir_path = os.path.join(artifacts_root_directory, "audiobooks", title)
            title_aax_path = os.path.join(title_dir_path, f"{title}.aax")
            title_m4b_path = os.path.join(title_dir_path, f"{title}.m4b")
            title_mp3_path = os.path.join(title_dir_path, f"{title}.mp3")

            # Load audiobook into AudioSegment so we can slice it
            audio_book = AudioSegment.from_mp3(
                title_mp3_path)

            file_counter = 1
            notes_dict = {}

            # Check whether a folder in clips/ for the book exists or not
            clips_dir_path = os.path.join(artifacts_root_directory, "audiobooks", title, "clips")
            path_exists = os.path.exists(clips_dir_path)
            if not path_exists:
                os.makedirs(clips_dir_path)

            for audio_clip in li_clips:
                # Get start position to slice
                raw_start_pos = int(audio_clip["startPosition"])

                # If we have a note then we save it so we can use it as the title for the bookmark text
                if audio_clip.get("type", None) in ["audible.note"]:
                    notes_dict[raw_start_pos] = audio_clip.get("text")
                    print(
                        f"CLIP: {notes_dict[raw_start_pos]}  {raw_start_pos}")

                if audio_clip.get("type", None) in ["audible.clip", "audible.bookmark"]:
                    start_pos = raw_start_pos - START_POSITION_OFFSET
                    end_pos = int(audio_clip.get(
                        "endPosition", raw_start_pos + 30000)) + END_POSITION_OFFSET
                    if start_pos == end_pos:
                        end_pos += 30000

                    # Slice it up
                    clip = audio_book[start_pos:end_pos]

                    file_name = notes_dict.get(
                        raw_start_pos, f"clip{file_counter}")

                    # Save the clip
                    clip_path = os.path.join(clips_dir_path, f"{file_name}.flac")
                    clip.export(
                        clip_path, format="flac")
                    file_counter += 1

    async def cmd_convert_audiobook(self):
        # FFMPEG needs to be installed for this step! see readme for more details
        li_books = await self.get_book_selection()

        for book in li_books:
            asin = book.get("asin")
            # Handle both string and nested dictionary formats for title
            title_value = book.get("title", {})
            if isinstance(title_value, str):
                _title = title_value
            else:
                _title = title_value.get("title", "untitled")
            
            if not _title:
                return

            title = _title.replace(" ", "_").lower()
            # Strips Audible DRM  from audiobook
            activation_bytes = self.get_activation_bytes()
            title_dir_path = os.path.join(artifacts_root_directory, "audiobooks", title)
            title_aax_path = os.path.join(title_dir_path, f"{title}.aax")
            title_m4b_path = os.path.join(title_dir_path, f"{title}.m4b")
            title_mp3_path = os.path.join(title_dir_path, f"{title}.mp3")
            os.system(
                f"ffmpeg -activation_bytes {activation_bytes} -i {title_aax_path} -c copy {title_m4b_path}")

            # Converts audiobook to .mp3
            os.system(
                f"ffmpeg -i {title_m4b_path} {title_mp3_path}")

    async def cmd_transcribe_bookmarks(self):
        li_books = await self.get_book_selection()

        r = sr.Recognizer()

        # Create dictionary to store titles and transcriptions and new folder to store transcriptions
        pairs = {}
        jsonHighlights = []
        
        for book in li_books:
            # Handle both string and nested dictionary formats for title
            title_value = book.get("title", {})
            if isinstance(title_value, str):
                _title = title_value
            else:
                _title = title_value.get("title", "untitled")
            
            # Handle authors similarly
            authors_value = book.get("title", {})
            if isinstance(authors_value, dict):
                _authors = authors_value.get("authors", [])
                if isinstance(_authors, list):
                    allAuthors = ", ".join(item.get('name', '') for item in _authors if isinstance(item, dict))
                else:
                    allAuthors = "Unknown Author"
            else:
                allAuthors = "Unknown Author"
            
            title = _title.lower().replace(" ", "_")
            title_dir_path = os.path.join(artifacts_root_directory, "audiobooks", title)
            clips_dir_path = os.path.join(title_dir_path, "clips")
            directory = os.fsencode(clips_dir_path)

            path_exists = os.path.exists(directory)
            if not path_exists:
                os.makedirs(directory)

            transcribed_clips_dir_path = os.path.join(title_dir_path, "trancribed_clips")
            trancribed_clips_path_exists = os.path.exists(transcribed_clips_dir_path)
            if not trancribed_clips_path_exists:
                os.makedirs(transcribed_clips_dir_path)

            for file in os.listdir(directory):
                highlight = {}
                filename = os.fsdecode(file)
                highlight["title"] = _title
                highlight["author"] = allAuthors
                if not filename.startswith("clip"):
                    highlight["note"] = filename.replace(".flac", "")
                highlight["source_type"] = "audible_bookmark_extractor"
                if filename.endswith(".flac"):
                    print(os.path.join(os.fsdecode(directory), filename))
                    heading = filename.replace(".flac", "")

                    audioclip = sr.AudioFile(os.path.join(
                        os.fsdecode(directory), filename))
                    with audioclip as source:
                        audio = r.record(source)

                    try:
                        text = r.recognize_google(audio)
                        pairs[str(heading)] = text
                        highlight["text"] = text
                    except Exception as e:
                        highlight["text"] = ""
                        print(f"Error while recognizing this clip {heading}: {e}")
                    xcel = pd.DataFrame(pairs.values(), index=pairs.keys())

                    # Change header format so that rows can be edited
                    pandas.io.formats.excel.ExcelFormatter.header_style = None

                    if highlight["text"]:
                        jsonHighlights.append(highlight)
                    
                    # Create writer instance with desired path
                    all_transcriptions_path = os.path.join(transcribed_clips_dir_path, "All_Transcriptions.xlsx")
                    writer = pd.ExcelWriter(
                        all_transcriptions_path, engine='xlsxwriter')

                    # Create a sheet in the same workbook for each file in the directory
                    sheet_name = title[:31].replace(":", "").replace("?", "")
                    xcel.to_excel(writer, sheet_name=sheet_name)
                    workbook = writer.book
                    worksheet = writer.sheets[sheet_name]

                    # Create header format to be used in all headers
                    header_format = workbook.add_format({
                        "valign": "vcenter",
                        "align": "center",
                        "bg_color": "#FFA500",
                        "bold": True,
                        "font_color": "#FFFFFF"})  # transcribe_bookmarks

                    # Set desired cell format
                    cell_format = workbook.add_format()
                    cell_format.set_align("vcenter")
                    cell_format.set_align("center")
                    cell_format.set_text_wrap(True)

                    # Apply header format and format columns to fit data
                    worksheet.write(0, 0, 'Clip Note', header_format)
                    worksheet.write(0, 1, 'Transcription', header_format)
                    worksheet.set_column("B:B", 100)
                    worksheet.set_column("A:A", 50)

                    # Format cells for appropiate size, wrap the text for style points
                    for i in range(1, (len(xcel)+1)):
                        worksheet.set_row(i, 100, cell_format)

                    # Apply changes and save xlsx to Transcribed bookmarks folder.
                    writer.close()
            transcription_contents_path = os.path.join(transcribed_clips_dir_path, "contents.json")
            with open(transcription_contents_path, "w") as f:
                json.dump(jsonHighlights, f, indent=4)                

    def get_activation_bytes(self):

        activation_bytes_path = os.path.join(artifacts_root_directory, "secrets", "activation_bytes.txt")
        # we already have activation bytes
        if os.path.exists(activation_bytes_path):
            with open(activation_bytes_path) as f:
                activation_bytes = f.readlines()[0]

        # we don't, so let's get them
        else:
            activation_bytes = self.auth.get_activation_bytes(
                activation_bytes_path, True)
            text_file = open(activation_bytes_path, "w")
            n = text_file.write(activation_bytes)
            text_file.close()

        return activation_bytes

    def bookmark_response_callback(self, resp):
        return resp

    async def cmd_export_bookmarks(self):
        """Export bookmarks to JSON file in current directory"""
        li_books = await self.get_book_selection()
        
        all_bookmarks = []
        
        for book in li_books:
            asin = book.get("asin")
            # Handle both string and nested dictionary formats for title
            title_value = book.get("title", {})
            if isinstance(title_value, str):
                _title = title_value
            else:
                _title = title_value.get("title", "untitled")
            
            if not _title:
                continue

            title = _title.lower().replace(" ", "_")
            
            print(f"Getting bookmarks for {_title}")
            
            # Get bookmarks from Audible API
            bookmarks_url = f"https://cde-ta-g7g.amazon.com/FionaCDEServiceEngine/sidecar?type=AUDI&key={asin}"
            
            try:
                with audible.Client(auth=self.auth, response_callback=self.bookmark_response_callback) as client:
                    library = client.get(
                        bookmarks_url,
                        num_results=1000,
                        response_groups="product_desc, product_attrs"
                    )
                    
                    li_bookmarks = library.json().get("payload", {}).get("records", [])
                    
                    # Process bookmarks into a simple format
                    for bookmark in li_bookmarks:
                        bookmark_data = {
                            "book_title": _title,
                            "asin": asin,
                            "type": bookmark.get("type", ""),
                            "start_position": bookmark.get("startPosition", 0),
                            "end_position": bookmark.get("endPosition", bookmark.get("startPosition", 0) + 30000),
                            "text": bookmark.get("text", ""),
                            "note": bookmark.get("note", ""),
                            "creation_time": bookmark.get("creationTime", "")
                        }
                        all_bookmarks.append(bookmark_data)
                        
            except Exception as e:
                print(f"Error getting bookmarks for {_title}: {e}")
                continue
        
        # Save to bookmarks.json in current directory
        import json
        import os
        
        output_file = "bookmarks.json"
        with open(output_file, 'w') as f:
            json.dump(all_bookmarks, f, indent=2)
        
        print(f"Bookmarks exported to: {output_file}")
        print(f"Total bookmarks exported: {len(all_bookmarks)}")
        
        # Verify the file was created and has content
        if os.path.exists(output_file) and os.path.getsize(output_file) > 0:
            print(f"✅ Export successful! File size: {os.path.getsize(output_file)} bytes")
        else:
            print("❌ Export failed - file was not created or is empty")

    async def cmd_export_bookmarks_simple(self, book_index=0):
        """Export bookmarks to JSON file in current directory with automatic book selection"""
        # Get library if not already loaded
        if not self.library:
            await self.get_library()
        
        # Select book by index (default to first book)
        try:
            book_index = int(book_index)
            if 0 <= book_index < len(self.library["items"]):
                selected_book = self.library["items"][book_index]
                li_books = [{"title": selected_book.get("title", 'untitled'), "asin": selected_book["asin"]}]
                print(f"Selected book: {selected_book.get('title', 'untitled')}")
            else:
                print(f"Invalid book index {book_index}. Available books: 0-{len(self.library['items'])-1}")
                return
        except (ValueError, IndexError):
            print("Invalid book index")
            return
        
        all_bookmarks = []
        
        for book in li_books:
            asin = book.get("asin")
            _title = book.get("title", 'untitled')
            if not _title:
                continue

            print(f"Getting bookmarks for {_title}")
            
            # Get bookmarks from Audible API
            bookmarks_url = f"https://cde-ta-g7g.amazon.com/FionaCDEServiceEngine/sidecar?type=AUDI&key={asin}"
            
            try:
                with audible.Client(auth=self.auth, response_callback=self.bookmark_response_callback) as client:
                    library = client.get(
                        bookmarks_url,
                        num_results=1000,
                        response_groups="product_desc, product_attrs"
                    )
                    
                    li_bookmarks = library.json().get("payload", {}).get("records", [])
                    
                    # Process bookmarks into a simple format
                    for bookmark in li_bookmarks:
                        # Convert milliseconds to a more standard format for the pipeline
                        start_ms = int(bookmark.get("startPosition", 0))
                        end_ms = int(bookmark.get("endPosition", start_ms + 30000))
                        
                        bookmark_data = {
                            "start_ms": start_ms,
                            "end_ms": end_ms,
                            "start": start_ms,  # Alternative field name
                            "end": end_ms,      # Alternative field name
                            "position": start_ms,  # Another alternative
                            "book_title": _title,
                            "asin": asin,
                            "type": bookmark.get("type", ""),
                            "text": bookmark.get("text", ""),
                            "note": bookmark.get("note", ""),
                            "creation_time": bookmark.get("creationTime", "")
                        }
                        all_bookmarks.append(bookmark_data)
                        
            except Exception as e:
                print(f"Error getting bookmarks for {_title}: {e}")
                continue
        
        # Save to bookmarks.json in current directory
        import json
        import os
        
        output_file = "bookmarks.json"
        with open(output_file, 'w') as f:
            json.dump(all_bookmarks, f, indent=2)
        
        print(f"Bookmarks exported to: {output_file}")
        print(f"Total bookmarks exported: {len(all_bookmarks)}")
        
        # Verify the file was created and has content
        if os.path.exists(output_file) and os.path.getsize(output_file) > 0:
            print(f"✅ Export successful! File size: {os.path.getsize(output_file)} bytes")
            return True
        else:
            print("❌ Export failed - file was not created or is empty")
            return False
