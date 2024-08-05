import requests

SEARX_URL = 'http://192.168.88.113:8888'
def search_searx(query, engines=None):
    search_url = f"{SEARX_URL}/search"
    params = {
        'q': query,  
        'format': 'json'  # 使用json响应格式
    }
    if engines:
        params['engines'] = ','.join(engines)  # 
    try:
        response = requests.get(search_url, params=params)
        response.raise_for_status()
        return response.json()  # 返回 JSON 数据
    except requests.RequestException as e:
        print(f"searX搜索失败: {e}")
        return None

def print_search_results(results):
    if results and 'results' in results:
        print(f"搜索结果: {results['results']}")
        for result in results['results']:
            title = result.get('title', '无标题')
            url = result.get('url', '无链接')
            snippet = result.get('content', '无摘要')
            print(f"标题: {title}")
            print(f"链接: {url}")
            print(f"摘要: {snippet}")
            print('-' * 80)
    else:
        print("没有搜索结果")

if __name__ == '__main__':
    query = '中通服设计'
    engines = ['baidu', 'sogou']

    results = search_searx(query, engines)
    print_search_results(results)
    # print(results)