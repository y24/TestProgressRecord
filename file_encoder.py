import sys
import base64

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