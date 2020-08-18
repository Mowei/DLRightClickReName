RJ改名工具
======


# 安裝
Install.bat

# 使用
右鍵
<a href="https://i.imgur.com/w6oNoqp.png"><img src="https://i.imgur.com/w6oNoqp.png" title="source: imgur.com" /></a>

# 設定
App.config     

NameFormatTemplate : 自定義重新命名檔名規則     
MakerNameXPath  :社團名稱     
WorkNameXPath   :標題名稱     
SaleDateXPath   :販售日     
WorkGenreXPath  :作品形式     
     
value 部分用html tag 去鎖定資料     

```=xml
<?xml version="1.0" encoding="utf-8"?>
<configuration>
  <configSections>
    <section name="genreTypes" type="System.Configuration.NameValueSectionHandler" />
  </configSections>
  <appSettings>
    <!--add key="NameFormatTemplate" value="[%maker_name%][%sale_date%][%number%]%work_name%%work_genre%" /-->
    <add key="NameFormatTemplate" value="[%maker_name%][%sale_date%][%number%]%work_name%%custom_types%" />
    <add key="WorkNameXPath" value="//h1[@id='work_name']//a" />
    <add key="MakerNameXPath" value="//span[@class='maker_name']" />
    <add key="SaleDateXPath" value="//table[@id='work_outline']//tr[1]//td" />
    <add key="WorkGenreXPath" value="//div[@class='work_genre']//span" />
    <!--Image-->
    <add key="ProductSampleImagesXPath" value="//div[@class='product-slider-data']//div" />
  </appSettings>
  <!--Types-->
  <genreTypes>
    <add key="TypeXPath1" value="//div[@class='work_genre']//span[@class='icon_SOU']" />
    <add key="TypeXPath2" value="//div[@class='work_genre']//span[@class='icon_MOV']" />
    <add key="TypeXPath3" value="//div[@class='work_genre']//span[@class='icon_MNG']" />
  </genreTypes>
</configuration>
```
