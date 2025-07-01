import os
import sys
import json
import socket
import ssl
import platform
import datetime
import subprocess
from pathlib import Path
from bs4 import BeautifulSoup
import requests
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry
    
class PracticeMojoTester:
    def __init__(self):
        self.results = []
        self.has_errors = False
        self.log_file = f"pmojo_test_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.log"
        
    def log(self, message, status="INFO"):
        """Log message to console and file"""
        timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        log_entry = f"[{timestamp}] [{status}] {message}"
        print(log_entry)
        
        with open(self.log_file, "a", encoding="utf-8") as f:
            f.write(log_entry + "\n")
            
        self.results.append((status, message))
        if status == "ERROR":
            self.has_errors = True
    
    def test_system_info(self):
        """Gather system information"""
        self.log("=== SYSTEM INFORMATION ===", "HEADER")
        self.log(f"Operating System: {platform.system()} {platform.release()}")
        self.log(f"Python Version: {sys.version}")
        self.log(f"Current Directory: {os.getcwd()}")
        
        # Check Python packages
        packages = {
            "requests": "requests",
            "bs4": "beautifulsoup4",
            "keyboard": "keyboard",
            "pywinauto": "pywinauto",
            "ahk": "ahk"
        }
        self.log("\nInstalled Packages:")
        for pkg in packages:
            try:
                __import__(pkg)
                self.log(f"{pkg} is installed", "SUCCESS")
            except ImportError:
                self.log(f"{pkg} is NOT installed", "ERROR")
    
    def test_network_basics(self):
        """Test basic network connectivity"""
        self.log("\n=== BASIC NETWORK TESTS ===", "HEADER")
        
        # Test internet connectivity
        try:
            socket.create_connection(("8.8.8.8", 53), timeout=3)
            self.log("Internet connection is working", "SUCCESS")
        except Exception as e:
            self.log(f"No internet connection: {e}", "ERROR")
            return
        
        # Test DNS resolution
        try:
            ip = socket.gethostbyname("app.practicemojo.com")
            self.log(f"DNS resolution successful: app.practicemojo.com -> {ip}", "SUCCESS")
        except Exception as e:
            self.log(f"DNS resolution failed: {e}", "ERROR")
            return
    
    def test_proxy_settings(self):
        """Check for proxy configurations"""
        self.log("\n=== PROXY SETTINGS ===", "HEADER")
        
        proxy_vars = ['HTTP_PROXY', 'HTTPS_PROXY', 'http_proxy', 'https_proxy', 
                     'ALL_PROXY', 'NO_PROXY']
        proxies_found = False
        
        for var in proxy_vars:
            if var in os.environ:
                self.log(f"{var}: {os.environ[var]}", "INFO")
                proxies_found = True
        
        if not proxies_found:
            self.log("No proxy environment variables found", "INFO")
        
        # Check Windows proxy settings
        if platform.system() == "Windows":
            try:
                import winreg
                with winreg.OpenKey(winreg.HKEY_CURRENT_USER,
                                  r"Software\Microsoft\Windows\CurrentVersion\Internet Settings") as key:
                    proxy_enable, _ = winreg.QueryValueEx(key, "ProxyEnable")
                    if proxy_enable:
                        proxy_server, _ = winreg.QueryValueEx(key, "ProxyServer")
                        self.log(f"Windows Proxy Enabled: {proxy_server}", "WARNING")
                    else:
                        self.log("Windows Proxy: Disabled", "INFO")
            except Exception as e:
                self.log(f"Could not check Windows proxy settings: {e}", "WARNING")
    
    def test_ssl_connection(self):
        """Test SSL/TLS connection"""
        self.log("\n=== SSL/TLS CONNECTION TEST ===", "HEADER")
        
        try:
            context = ssl.create_default_context()
            with socket.create_connection(("app.practicemojo.com", 443), timeout=10) as sock:
                with context.wrap_socket(sock, server_hostname="app.practicemojo.com") as ssock:
                    self.log(f"SSL connection successful", "SUCCESS")
                    self.log(f"  Protocol: {ssock.version()}")
                    self.log(f"  Cipher: {ssock.cipher()}")
                    
                    # Get certificate info
                    cert = ssock.getpeercert()
                    if cert:
                        self.log(f"  Certificate Subject: {cert.get('subject', 'N/A')}")
                        self.log(f"  Certificate Issuer: {cert.get('issuer', 'N/A')}")
        except ssl.SSLError as e:
            self.log(f"SSL Error (possibly corporate firewall): {e}", "ERROR")
        except Exception as e:
            self.log(f"SSL connection failed: {e}", "ERROR")
    
    def test_http_requests(self):
        """Test HTTP requests with various methods"""
        self.log("\n=== HTTP REQUEST TESTS ===", "HEADER")
        
        # Test basic requests
        test_urls = [
            ("https://httpbin.org/get", "Test site"),
            ("https://app.practicemojo.com", "PracticeMojo homepage"),
            ("https://app.practicemojo.com/Pages/login", "PracticeMojo login page")
        ]
        
        session = requests.Session()
        
        for url, description in test_urls:
            try:
                self.log(f"\nTesting {description}: {url}")
                resp = session.get(url, timeout=10)
                self.log(f"  Status Code: {resp.status_code}", "SUCCESS")
                self.log(f"  Response Length: {len(resp.text)} bytes")
                self.log(f"  Headers: {dict(resp.headers)}")
                
                if "practicemojo" in url and resp.status_code == 200:
                    self.log("  PracticeMojo is accessible", "SUCCESS")
                    
            except requests.exceptions.SSLError as e:
                self.log(f"  SSL Error: {e}", "ERROR")
                self.log("  This often indicates a corporate firewall intercepting SSL", "WARNING")
                
            except requests.exceptions.ConnectionError as e:
                self.log(f"  Connection Error: {e}", "ERROR")
                
            except requests.exceptions.Timeout:
                self.log(f"  Request timed out (>10 seconds)", "ERROR")
                
            except Exception as e:
                self.log(f"  Request failed: {e}", "ERROR")
    
    def test_practicemojo_login(self):
        """Test actual PracticeMojo login process"""
        self.log("\n=== PRACTICEMOJO LOGIN TEST ===", "HEADER")
        
        # Check for config file
        config_path = Path("config.json")
        if not config_path.exists():
            self.log("config.json not found!", "ERROR")
            self.log("  Please create config.json with USERNAME and PASSWORD", "ERROR")
            return
        
        try:
            with open(config_path) as f:
                config = json.load(f)
                username = config.get("USERNAME", "").strip()
                password = config.get("PASSWORD", "").strip()
                
            if not username or not password:
                self.log("Username or password is empty in config.json", "ERROR")
                return
                
            self.log("Config file loaded successfully", "SUCCESS")
            self.log(f"  Username: {username[:3]}***{username[-1:] if len(username) > 4 else ''}")
            
        except json.JSONDecodeError as e:
            self.log(f"config.json is not valid JSON: {e}", "ERROR")
            return
        except Exception as e:
            self.log(f"Error reading config: {e}", "ERROR")
            return
        
        # Create session with retry
        session = requests.Session()
        retry_strategy = Retry(
            total=3,
            backoff_factor=1,
            status_forcelist=[500, 502, 503, 504]
        )
        adapter = HTTPAdapter(max_retries=retry_strategy)
        session.mount("http://", adapter)
        session.mount("https://", adapter)
        
        # Test login process step by step
        self.log("\nStep 1: Fetching login page...")
        try:
            login_page_url = "https://app.practicemojo.com/Pages/login"
            resp = session.get(login_page_url, timeout=30)
            self.log(f"  Login page fetched (status: {resp.status_code})", "SUCCESS")
            
            # Check if we got a valid login page
            if "login" in resp.text.lower() or "password" in resp.text.lower():
                self.log("  Login form detected", "SUCCESS")
            else:
                self.log("  Login form not detected in response", "WARNING")
                
        except Exception as e:
            self.log(f"  Failed to fetch login page: {e}", "ERROR")
            return
        
        self.log("\nStep 2: Attempting login...")
        try:
            login_url = "https://app.practicemojo.com/cgi-bin/WebObjects/PracticeMojo.woa/wa/userLogin"
            form_data = {
                "loginId": username,
                "password": password,
                "slug": "login"
            }
            headers = {
                "Referer": login_page_url,
                "Content-Type": "application/x-www-form-urlencoded",
                "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"
            }
            
            resp = session.post(login_url, data=form_data, headers=headers, timeout=30)
            self.log(f"  Login response status: {resp.status_code}")
            
            if resp.status_code == 200:
                if "Invalid login" in resp.text or "Incorrect password" in resp.text:
                    self.log("  Login failed: Invalid credentials", "ERROR")
                elif "dashboard" in resp.text.lower() or "logout" in resp.text.lower():
                    self.log("  Login appears successful!", "SUCCESS")
                else:
                    self.log("  Login status unclear, response saved to login_response.html", "WARNING")
                    with open("login_response.html", "w", encoding="utf-8") as f:
                        f.write(resp.text)
            else:
                self.log(f"  Unexpected status code: {resp.status_code}", "ERROR")
                
        except Exception as e:
            self.log(f"  Login request failed: {e}", "ERROR")
    
    def test_firewall_and_antivirus(self):
        """Check for common firewall/AV software"""
        self.log("\n=== SECURITY SOFTWARE CHECK ===", "HEADER")
        
        if platform.system() == "Windows":
            # Check Windows Defender Firewall
            try:
                result = subprocess.run(
                    ["netsh", "advfirewall", "show", "currentprofile"], 
                    capture_output=True, text=True
                )
                if "State" in result.stdout and "ON" in result.stdout:
                    self.log("Windows Firewall is ENABLED", "WARNING")
                else:
                    self.log("Windows Firewall is disabled or not detected", "INFO")
            except:
                self.log("Could not check Windows Firewall status", "INFO")
            
            # Check for common AV processes
            av_processes = {
                "MsMpEng.exe": "Windows Defender",
                "avast": "Avast",
                "avg": "AVG", 
                "mcafee": "McAfee",
                "norton": "Norton",
                "kaspersky": "Kaspersky",
                "bitdefender": "Bitdefender",
                "malwarebytes": "Malwarebytes"
            }
            
            try:
                processes = subprocess.check_output("tasklist", shell=True, text=True).lower()
                for proc, name in av_processes.items():
                    if proc.lower() in processes:
                        self.log(f"{name} detected - may interfere with connections", "WARNING")
            except:
                self.log("Could not check for antivirus software", "INFO")
    
    def generate_report(self):
        """Generate final report"""
        self.log("\n=== TEST SUMMARY ===", "HEADER")
        
        error_count = sum(1 for status, _ in self.results if status == "ERROR")
        warning_count = sum(1 for status, _ in self.results if status == "WARNING")
        success_count = sum(1 for status, _ in self.results if status == "SUCCESS")
        
        self.log(f"Tests completed: {len(self.results)}")
        self.log(f"Successes: {success_count}")
        self.log(f"Warnings: {warning_count}")
        self.log(f"Errors: {error_count}")
        
        if self.has_errors:
            self.log("\nISSUES DETECTED - Please review the errors above", "WARNING")
            self.log(f"Full log saved to: {self.log_file}", "INFO")
        else:
            self.log("\nAll tests passed!", "SUCCESS")
        
        # Provide recommendations
        self.log("\n=== RECOMMENDATIONS ===", "HEADER")
        
        for status, message in self.results:
            if status == "ERROR":
                if "SSL" in message:
                    self.log("SSL issues detected - try disabling SSL verification or check with IT about firewall")
                elif "proxy" in message.lower():
                    self.log("Proxy issues detected - check with IT for correct proxy settings")
                elif "DNS" in message:
                    self.log("DNS issues detected - try using Google DNS (8.8.8.8)")
                elif "credentials" in message.lower():
                    self.log("Login credentials appear to be incorrect - verify config.json")


def main():
    print("=" * 60)
    print("PracticeMojo Connection Tester")
    print("=" * 60)
    print()
    
    tester = PracticeMojoTester()
    
    try:
        tester.test_system_info()
        tester.test_network_basics()
        tester.test_proxy_settings()
        tester.test_ssl_connection()
        tester.test_http_requests()
        tester.test_firewall_and_antivirus()
        tester.test_practicemojo_login()
    except KeyboardInterrupt:
        print("\n\nTest interrupted by user")
    except Exception as e:
        tester.log(f"Unexpected error during testing: {e}", "ERROR")
    finally:
        tester.generate_report()
    
    print(f"\n\nDetailed log saved to: {tester.log_file}")
    print("Please share this log file when reporting issues.")
    input("\nPress Enter to exit...")


if __name__ == "__main__":
    main()