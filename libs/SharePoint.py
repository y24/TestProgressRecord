from urllib.parse import urlparse

def generate_download_url(share_link: str, site_url: str) -> str:
    """
    共有リンクURLとサイトURLから、SharePointのダウンロード用URLを生成する。

    Parameters:
        share_link (str): 共有リンクのURL
        site_url (str): サイトのベースURL

    Returns:
        str: ダウンロード用URL
    """
    # クエリ部分を取り除く
    parsed_url = urlparse(share_link)
    path = parsed_url.path

    # 最後のスラッシュ以降の部分を取得
    share_id = path.split('/')[-1]

    # ダウンロードURLの組み立て
    download_url = f"{site_url}/_layouts/download.aspx?share={share_id}"
    return download_url
