import json, os, sys, tempfile, urllib.parse, requests
from requests_ntlm import HttpNtlmAuth
from typing import List

def is_sharepoint_url(url: str) -> bool:
    """URLがSharePointのURLかどうかを判定する"""
    parsed_url = urllib.parse.urlparse(url)
    return "sharepoint.com" in parsed_url.netloc

def resolve_sharepoint_link(shared_link):
    try:
        response = requests.head(shared_link, allow_redirects=False)
        if response.status_code in (301, 302):
            # リダイレクト先URLを取得
            redirect_url = response.headers.get('Location')
            return redirect_url
        else:
            return shared_link  # リダイレクトがなければそのまま
    except Exception as e:
        raise Exception(f"リダイレクト解決中にエラーが発生しました: {e}")

def parse_sharepoint_link(resolved_url):
    parsed_url = urllib.parse.urlparse(resolved_url)
    path = parsed_url.path  # ex) /sites/ProjectX/Shared%20Documents/Design/Spec.docx

    # URLパターンを解析
    if path.startswith('/sites/'):
        # /sites/{site}/... の形
        parts = path.split('/')
        if len(parts) >= 3:
            tenant = parsed_url.netloc  # contoso.sharepoint.com
            site = parts[2]             # ProjectX
            rest_path = '/'.join(parts[3:])  # Shared Documents/Design/Spec.docx

            site_url = f"https://{tenant}/sites/{site}"
            file_path = f"/sites/{site}/{rest_path}"
            # URLデコード
            file_path = urllib.parse.unquote(file_path)

            return site_url, file_path
    raise Exception("解析できないURL形式です。")


def download_sharepoint_file(url: str, temp_dir: str) -> str:
    """SharePointのファイルをダウンロードする"""
    try:
        # Windows認証を使用
        response = requests.get(url, stream=True, auth=HttpNtlmAuth())
        response.raise_for_status()
        
        # URLからファイル名を取得
        file_name = os.path.basename(urllib.parse.urlparse(url).path)
        if not file_name:
            file_name = "downloaded_file"
            
        # 一時フォルダにファイルを保存
        file_path = os.path.join(temp_dir, file_name)
        with open(file_path, 'wb') as f:
            for chunk in response.iter_content(chunk_size=8192):
                if chunk:
                    f.write(chunk)
                    
        return file_path
    except Exception as e:
        raise Exception(f"ファイルのダウンロードに失敗しました。\n{str(e)}")

def process_json_file(json_path: str) -> List[str]:
    """
    JSONファイルからURLを取得し、ファイルをダウンロードする
    
    Args:
        json_path (str): JSONファイルのパス
        
    Returns:
        List[str]: ダウンロードされたファイルのパスのリスト
    """
    try:
        # JSONファイルを読み込む
        with open(json_path, 'r', encoding='utf-8') as f:
            data = json.load(f)
            
        # projectオブジェクトのfilesリストからURLを取得
        urls = []
        if 'project' in data and 'files' in data['project']:
            urls = [file.get('url') for file in data['project']['files'] if file.get('url')]
            
        # 一時フォルダを作成
        with tempfile.TemporaryDirectory(prefix="_TEMP_") as temp_dir:
            downloaded_files = []
            
            # 各URLに対して処理を実行
            for url in urls:
                if is_sharepoint_url(url):
                    file_path = download_sharepoint_file(url, temp_dir)
                    downloaded_files.append(file_path)
                    
            return downloaded_files, temp_dir
            
    except Exception as e:
        raise Exception(str(e))


def main():
    if len(sys.argv) != 2:
        print("使い方: python script.py <共有リンクURL>")
        sys.exit(1)

    shared_link = sys.argv[1]
    resolved_url = resolve_sharepoint_link(shared_link)
    site_url, file_path = parse_sharepoint_link(resolved_url)

    print(f"サイトURL: {site_url}")
    print(f"ファイルパス: {file_path}")

    if __name__ == "__main__":
        main()