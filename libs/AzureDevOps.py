import requests
from datetime import datetime
from collections import defaultdict
from typing import Dict, List, Tuple
from .Win32Credential import load_api_key

class AzureDevOpsAPI:
    def __init__(self):
        self.api_key = load_api_key()
        if not self.api_key:
            raise ValueError("APIキーが設定されていません。")
        
        self.base_url = "https://dev.azure.com/{organization}/{project}/_apis"
        self.headers = {
            'Authorization': f'Basic {self.api_key}',
            'Content-Type': 'application/json'
        }

    def get_work_item_children(self, work_item_id: int) -> List[Dict]:
        """指定されたWorkItemの子要素を取得する"""
        url = f"{self.base_url}/wit/workitems/{work_item_id}?$expand=relations&api-version=7.0"
        response = requests.get(url, headers=self.headers)
        response.raise_for_status()
        
        work_item = response.json()
        children = []
        
        for relation in work_item.get('relations', []):
            if relation['rel'] == 'System.LinkTypes.Hierarchy-Forward':
                child_id = int(relation['url'].split('/')[-1])
                child_url = f"{self.base_url}/wit/workitems/{child_id}?api-version=7.0"
                child_response = requests.get(child_url, headers=self.headers)
                child_response.raise_for_status()
                children.append(child_response.json())
        
        return children

    def analyze_work_items(self, work_item_id: int, target_type: str) -> Dict[Tuple[str, str], int]:
        """WorkItemの子要素を分析し、作成者と日付で集計する"""
        children = self.get_work_item_children(work_item_id)
        filtered_children = [child for child in children if child['fields']['System.WorkItemType'] == target_type]
        
        result = defaultdict(int)
        
        for child in filtered_children:
            created_by = child['fields']['System.CreatedBy']['displayName']
            created_date = datetime.strptime(
                child['fields']['System.CreatedDate'],
                '%Y-%m-%dT%H:%M:%S.%fZ'
            ).strftime('%Y-%m-%d')
            
            result[(created_by, created_date)] += 1
        
        return dict(result)

def analyze_work_item_children(work_item_id: int, work_item_type: str) -> Dict[Tuple[str, str], int]:
    """メインの分析関数"""
    api = AzureDevOpsAPI()
    return api.analyze_work_items(work_item_id, work_item_type)

if __name__ == "__main__":
    print(analyze_work_item_children(1, "Bug"))
