
import json
import os
import argparse
from google.oauth2 import service_account
from googleapiclient.discovery import build

def get_drive_service(creds_path):
    scopes = ['https://www.googleapis.com/auth/drive']
    creds = service_account.Credentials.from_service_account_file(
        creds_path, scopes=scopes)
    service = build('drive', 'v3', credentials=creds)
    return service

def check_quota(service):
    try:
        about = service.about().get(fields="user,storageQuota").execute()
        user_info = about.get('user', {})
        quota = about.get('storageQuota', {})
        
        email = user_info.get('emailAddress', 'Unknown')
        usage = int(quota.get('usage', 0))
        limit_str = quota.get('limit')
        limit = int(limit_str) if limit_str else -1
        
        usage_mb = usage / (1024 * 1024)
        limit_mb = limit / (1024 * 1024) if limit != -1 else -1
        limit_disp = f"{limit_mb:.2f} MB" if limit != -1 else "Unlimited"
        
        print(f"--- Service Account Info ---")
        print(f"Email: {email}")
        print(f"Usage: {usage_mb:.2f} MB")
        print(f"Limit: {limit_disp}")
        
    except Exception as e:
        print(f"Error checking quota: {e}")

if __name__ == "__main__":
    # Look for likely json files in current directory
    json_files = [f for f in os.listdir('.') if f.endswith('.json') and 'service' in f.lower()]
    
    if not json_files:
        print("No service_account.json found in current directory.")
        exit(1)
        
    print(f"Found credential file: {json_files[0]}")
    service = get_drive_service(json_files[0])
    check_quota(service)
