import sys
import base64

def create_download_link(b64_content: str, filename: str, link_text: str = None) -> str:
    """
    Base64エンコードされたデータを含むダウンロードリンクのHTML文字列を生成します。
    
    Args:
        b64_content (str): Base64エンコードされたファイルの内容
        filename (str): ダウンロード時のファイル名
        link_text (str, optional): リンクに表示するテキスト。指定がない場合はファイル名が使用されます。
    
    Returns:
        str: HTML形式のダウンロードリンク文字列
    """
    link_text = link_text or filename
    return f'<a href="data:application/octet-stream;base64,{b64_content}" download="{filename}" style="text-decoration:none;">{link_text}</a>'

def encode_file(file_path):
    try:
        with open(file_path, "rb") as f:
            content = f.read()
        encoded = base64.b64encode(content).decode()
        print(encoded)  # 標準出力に結果を出力
        return 0
    except Exception as e:
        print(f"Error: {str(e)}", file=sys.stderr)
        return 1

if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("Usage: python file_encoder.py <file_path>", file=sys.stderr)
        sys.exit(1)
    
    file_path = sys.argv[1]
    sys.exit(encode_file(file_path)) 