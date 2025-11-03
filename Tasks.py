import json
import os
import sys
import subprocess
import threading
import queue
import tkinter as tk
from tkinter import ttk, scrolledtext
from datetime import datetime
import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from playwright.sync_api import sync_playwright, Playwright, TimeoutError, expect
import warnings
import pyxlsb
import csv

import time
from datetime import date, timedelta


warnings.filterwarnings("ignore", category=UserWarning)

import requests
from playwright.sync_api import Page, TimeoutError




def wait_for_server_ready(q, url="http://localhost:5000", retries=5, delay=2):
    """Polls the API endpoint until it responds."""
    q.put(("status", "⏳ Waiting for Captcha API server to start..."))
    for i in range(retries):
        try:
            # Try connecting to the root URL (or a health check endpoint if available)
            response = requests.get(url, timeout=10) 
            if response.status_code < 500: # Any client or server success code
                q.put(("status", "API server is ready."))
                return True
        except requests.exceptions.ConnectionError:
            q.put(("status", f"Waiting for API... Attempt {i+1}/{retries}"))
            time.sleep(delay)
    return False

def resolve_captcha(page: Page, q) -> str | None:
    # 🎯 Locator for the captcha image
    captcha_locator = page.locator('#imgCaptcha')

    try:
        # Wait for captcha to be visible and stable
        captcha_locator.wait_for(state="visible", timeout=10000)

        # Get the captcha image as bytes (exact one shown on screen)
        q.put(("status", "📸 Capturing CAPTCHA image from page..."))
        image_bytes = captcha_locator.screenshot()  # <-- this captures the exact image shown

        # Save locally for debugging/logging
        with open("captcha.png", "wb") as f:
            f.write(image_bytes)

    except TimeoutError:
        q.put(("status", "❌ Captcha image not found or not visible."))
        return None
    except Exception as e:
        q.put(("status", f"❌ Error capturing captcha screenshot: {e}"))
        return None

    # --- Send Image to Local API Endpoint ---
    endpoint_url = "http://localhost:5000/resolver_captcha"
    q.put(("status", f"📤 Sending Captcha to local API: {endpoint_url}"))

    try:
        files = {'imagem': ('captcha.png', image_bytes, 'image/png')}
        api_response = requests.post(endpoint_url, files=files, timeout=20)
        api_response.raise_for_status()

        data = api_response.json()
        captcha_text = data.get("resultado", "").strip()

        if captcha_text:
            q.put(("status", f"✅ Captcha solved! Text: {captcha_text}"))
            return captcha_text
        else:
            q.put(("status", f"❌ Captcha API returned no text: {data}"))
            return None

    except requests.exceptions.RequestException as e:
        q.put(("status", f"❌ Error contacting Captcha API: {e}"))
        return None


def Login_and_Navigation(page: Page, url, q, username, password):
    # This function uses the new resolve_captcha helper
    try:
        # --- 1. Login and Initial Navigation ---
        q.put(("status", "Navigating to login page..."))
        page.goto(url, timeout=60000)
        api_process = start_apicapcha_server(q)

        if not api_process:
            q.put(("status", "🛑 Failed to start Captcha API server. Aborting."))
            return

        # ⏳ WAIT FOR THE SERVER TO BE READY
        if not wait_for_server_ready(q):
            stop_apicapcha_server(q)
            
            q.put(("status", "🛑 Captcha API server failed to become ready. Aborting."))
       
        q.put(("progress", 2))
        q.put(("status", "Performing login..."))
        page.get_by_role("textbox", name="E-mail").fill(username)
        
        


        MAX_RETRIES = 20  # optional safety limit

        for attempt in range(1, MAX_RETRIES + 1):
            try:
                # Wait for the post-login element to appear (login success indicator)
                page.wait_for_selector("text=Solicitar agendamentos de", timeout=1000)
                q.put(("status", f"✅ Login successful on attempt {attempt}!"))
                break  # ✅ Success — exit loop

            except TimeoutError:
                # ❌ Element not found in time — try login again
                q.put(("status", f"⏳ Attempt {attempt}: Element not found, retrying login..."))

                # Fill password again
                page.get_by_role("textbox", name="Senha").fill(password)

                # --- CAPTCHA RESOLUTION ---
                captcha_text = resolve_captcha(page, q)
                if not captcha_text:
                    q.put(("status", "🛑 Failed to resolve CAPTCHA. Aborting login."))
                    return  # Exit function if CAPTCHA failed

                page.get_by_role("textbox", name="Repita o texto da imagem").fill(captcha_text)
                page.get_by_role("button", name="ENTRAR").click()

                # Wait briefly before checking again
                time.sleep(2)

        else:
            # If loop ends without break
            q.put(("status", "❌ Login failed after maximum retries."))

        q.put(("status", "Login sucessful")) # Change to attempted until success is confirmed
        q.put(("progress", 5))
        
        


    except Exception as e:
        q.put(("status", f"An error occurred: {e}"))



_API_PROCESS = None 

def start_apicapcha_server(q):
    """Starts the Captcha API server as a background process using Popen."""
    global _API_PROCESS
    
    script_path = r"C:\Users\perna\Desktop\Barrueri\RPA- ATACADO\ApiCaptcha-main\main.py"
    script_dir = r"C:\Users\perna\Desktop\Barrueri\RPA- ATACADO\ApiCaptcha-main"
    
    if _API_PROCESS and _API_PROCESS.poll() is None:
        q.put(("status", "✅ ApiCaptcha server is already running."))
        return _API_PROCESS

    if not os.path.exists(script_path) or not os.path.isdir(script_dir):
        q.put(("status", "❌ ApiCaptcha script path error."))
        return None

    try:
        q.put(("status", "🚀 Starting ApiCaptcha server in background..."))
        
        # Use Popen for non-blocking execution
        process = subprocess.Popen(
            [sys.executable, script_path], 
            cwd=script_dir
        )
        _API_PROCESS = process
        q.put(("status", f"✅ ApiCaptcha server started with PID: {process.pid}"))
        return process 
    except Exception as e:
        q.put(("status", f"❌ Error starting API server: {e}"))
        return None

def stop_apicapcha_server(q):
    """Terminates the background server process."""
    global _API_PROCESS
    if _API_PROCESS and _API_PROCESS.poll() is None:
        q.put(("status", "🛑 Stopping ApiCaptcha server..."))
        _API_PROCESS.terminate()
        _API_PROCESS.wait(timeout=5)
        q.put(("status", "🛑 ApiCaptcha server stopped."))
    _API_PROCESS = None