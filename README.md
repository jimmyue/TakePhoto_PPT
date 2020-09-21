# TakePhoto_PPT

1.selenium驱动远程机访问系统，并截图

2.查询数据库动态数据

3.图片导入PPT自动生成报告

4.Jenkins集成，动态传递参数


#使用docker部署的selenium远程端执行
chrome={'browserName': 'chrome', 'maxInstances': '5', 'platform': 'LINUX', 'platformName': 'LINUX', 'seleniumProtocol': 'WebDriver', 'version': '84.0.4147.105'}
opera={'browserName': 'operablink', 'maxInstances': '5', 'platform': 'LINUX', 'platformName': 'LINUX', 'seleniumProtocol': 'WebDriver'}
firefox={'browserName': 'firefox', 'maxInstances': '5', 'platform': 'LINUX', 'platformName': 'LINUX', 'seleniumProtocol': 'WebDriver', 'version': '78.0.2'}
driver = webdriver.Remote(command_executor="http://10.10.10.71:4444/wd/hub",desired_capabilities=chrome) #修改浏览器
