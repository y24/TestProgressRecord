import win32cred
import win32con

TARGET_NAME = 'MyApp_APIKey'

def save_api_key(api_key: str):
    credential = {
        'Type': win32cred.CRED_TYPE_GENERIC,
        'TargetName': TARGET_NAME,
        'CredentialBlob': api_key.encode('utf-16'),
        'Persist': win32cred.CRED_PERSIST_ENTERPRISE,  # 管理者権限不要
        'UserName': 'APIUser'
    }
    win32cred.CredWrite(credential, 0)
    print("APIキーを保存しました。")

def load_api_key():
    try:
        cred = win32cred.CredRead(TARGET_NAME, win32cred.CRED_TYPE_GENERIC)
        return cred['CredentialBlob'].decode('utf-16')
    except Exception as e:
        print("読み込み失敗:", e)
        return None

def delete_api_key():
    try:
        win32cred.CredDelete(TARGET_NAME, win32cred.CRED_TYPE_GENERIC, 0)
        print("APIキーを削除しました。")
    except Exception as e:
        print("削除失敗:", e)
