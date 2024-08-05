import requests  
  
def search_with_searx(query, searx_url="http://192.168.88.113:8888"):  
    # 构造请求 URL，通常 Searx 允许你通过 URL 参数发送查询  
    search_url = f"{searx_url}/?q={query}"  
      
    # 发送 GET 请求  
    response = requests.get(search_url)  
      
    # 检查响应状态码  
    if response.status_code == 200:  
        # 返回响应内容，这里假定 Searx 返回的是 HTML  
        return response.text  
    else:  
        return f"Failed to retrieve results, status code: {response.status_code}"  
  
# 示例查询  
query = "python"  
results = search_with_searx(query)  
print(results)  # 这里将打印出 HTML 源代码，你可以解析它以获取结果