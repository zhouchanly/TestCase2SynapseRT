from playwright.sync_api import Playwright, sync_playwright
with sync_playwright() as p:
    # 可以选择chromium、firefox和webkit
    browser_type = p.chromium
    # 运行chrome浏览器，executablePath指定本地chrome安装路径
    # browser = browser_type.launch(headless=False,slowMo=50,executablePath=r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe")
    browser = browser_type.launch(headless=False)
    page = browser.new_page()
    page.goto('http://jira.intretech.com:8080/secure/Dashboard.jspa')
    # page.screenshot(path=f'example-{browser_type.name}.png')
    print(page.title())
    
    browser.close()