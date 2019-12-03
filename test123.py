import ActionClasses.Browser as b
import  ActionClasses.BrowserNav
import ActionClasses.Elements
import  ActionClasses.ClickElement
import  ActionClasses.SetText
import ActionClasses.CloseBrowser

result=b.execute("https://www.google.com")
#result1=ActionClasses.BrowserNav.BrowserNavigation().execute(Nav="backward",Url=None)
IDType="xpath"
ID="(//a)[11]"
result1=ActionClasses.ClickElement.ClickElement.execute(IDType,ID)
result2=ActionClasses.BrowserNav.BrowserNavigation.execute(Nav="backward",Url=None)
IDType1="name"
ID1="q"
text="hello"
result3=ActionClasses.SetText.SetText.execute(IDType1,ID1,text)
result4=ActionClasses.CloseBrowser.CloseBrowser.execute()


print(result,result1,result2,result3)


