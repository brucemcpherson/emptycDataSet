'gistThat@mcpher.com :do not modify this line - see ramblings.mcpher.com for details: updated on 8/18/2014 4:47:50 PM : from manifest:5055578 gist https://gist.github.com/brucemcpherson/3423885/raw/restLibrary.vba
Option Explicit
' v2.26
'for more about this
' http://ramblings.mcpher.com/Home/excelquirks/classeslink/data-manipulation-classes
'to contact me
' http://groups.google.com/group/excel-ramblings
'reuse of code
' http://ramblings.mcpher.com/Home/excelquirks/codeuse
' restlibrary - this is an automated rest query to excel table set of known queries
'
Const getItFrom = ""
''Const getItFrom = "https://script.google.com/a/macros/mcpher.com/s/AKfycbzLqpnQ2ey8CKAMmzchb2n2FU-aiae0iTKPzAOfAgEpxGwaJgk/exec"

' simplified interface
Public Function generalQuery(sheetName As String, _
                libEntry As String, queryString As String, _
                Optional breport As Boolean = True, _
                Optional queryCanBeBlank As Boolean = False, _
                Optional appendQuery As String = vbNullString) As cRest
    
        Set generalQuery = generalReport( _
            restQuery(sheetName, libEntry, queryString, , , , , , , , queryCanBeBlank, , , , , , appendQuery), breport)

End Function
Public Function generalDataSetQuery(sheetName As String, _
                libEntry As String, colName As String, _
                Optional breport As Boolean = True, _
                Optional queryCanBeBlank As Boolean = False, _
                Optional appendQuery As String = vbNullString, _
                Optional collectionNeeded As Boolean = True) As cRest

    Set generalDataSetQuery = generalReport( _
            restQuery(sheetName, libEntry, , colName, _
            , , , , , , , , , , , , appendQuery, collectionNeeded), breport)
    
End Function

Public Function generalReport(cr As cRest, breport As Boolean) As cRest
    If cr Is Nothing Then
        MsgBox ("failed to get any data")
    Else
        If breport Then
            MsgBox (cr.jObjects.count & " items retrieved ")
        End If
    End If
    Set generalReport = cr
End Function
Public Function getRestLibrary() As cJobject
    ' build it locally as previously
    Dim cb As cBrowser, cj As cJobject
    
    If getItFrom = vbNullString Then
        Set getRestLibrary = createRestLibrary
    Else
        ' get it from an API server
        Set cb = New cBrowser
        cb.init
        Set cj = New cJobject
        Set getRestLibrary = cj.init(Nothing).deSerialize(cb.httpGET(getItFrom))
    End If
End Function
Public Function createRestLibrary() As cJobject
    ' this creates the restlibrary as a jSon object
    Dim cj As cJobject
    Set cj = New cJobject
    cj.init Nothing, "restLibrary"

    With cj
    
        With .add("sunrise-sunset")
            .add "restType", erQueryPerRow
            .add "url", "http://api.sunrise-sunset.org/json?"
            .add "results", "results"
            .add "treeSearch", False
            .add "ignore", vbNullString
        End With

        With .add("lescourses")
            .add "restType", erSingleQuery
            .add "url", "http://www.geny.com/stats-records-hand-flux-donnees?typeStats=jockey-pmu&type=json&id_course="
            .add "results", "ResultSet.partants"
            .add "treeSearch", True
            .add "ignore", vbNullString
        End With
        With .add("imdb by id")
            .add "restType", erQueryPerRow
            .add "url", "http://www.omdbapi.com/?i="
            .add "results", vbNullString
            .add "treeSearch", False
            .add "ignore", vbNullString
        End With

        With .add("ua accounts")
            .add "restType", erSingleQuery
            .add "url", "https://www.googleapis.com/analytics/v3/management/accounts"
            .add "results", "items"
            .add "treeSearch", True
            .add "ignore", vbNullString
            .add "authType", erOAUTH2
            .add "authScope", "analytics"
        End With
        
        With .add("ua web properties")
            .add "restType", erSingleQuery
            .add "url", "https://www.googleapis.com/analytics/v3/management/accounts/"
            .add "results", "items"
            .add "treeSearch", True
            .add "ignore", vbNullString
            .add "authType", erOAUTH2
            .add "authScope", "analytics"
            .add "append", "/webproperties"
        End With
        
        With .add("ua data")
            .add "restType", erSingleQuery
            .add "url", "https://www.googleapis.com/analytics/v3/data/ga?ids=ga:"
            .add "results", ""
            .add "treeSearch", True
            .add "ignore", vbNullString
            .add "authType", erOAUTH2
            .add "authScope", "analytics"
        End With
        

        With .add("open weather xml")
            .add "restType", erQueryPerRow
            .add "url", "http://api.openweathermap.org/data/2.5/weather?q="
            .add "results", "current"
            .add "treeSearch", True
            .add "ignore", vbNullString
            .add "append", "&mode=xml"
            .add "resultsFormat", erAUTO
        End With
        
        With .add("funds")
            .add "restType", erSingleQuery
            .add "url", "https://newtemplate.hosts.webrecs.com/alfresco/service/webrecs/fundsearcher.xml?full=true"
            .add "results", "funds.sites"
            .add "treeSearch", True
            .add "ignore", vbNullString
            .add "resultsFormat", erAUTO
        End With
        
        With .add("fusiondata")
            .add "restType", erSingleQuery
            .add "url", "https://www.googleapis.com/fusiontables/v1/query?key="
            .add "results", ""
            .add "treeSearch", True
            .add "ignore", vbNullString
            .add "append", "&sql="
        End With
    
        With .add("my society")
            .add "restType", erQueryPerRow
            .add "url", "http://mapit.mysociety.org/postcode/"
            .add "results", "areas"
            .add "treeSearch", True
            .add "ignore", vbNullString
            .add "alwaysEncode", True
        End With
        
        With .add("tagsite")
            .add "restType", erSingleQuery
            .add "url", "https://script.google.com/macros/s/AKfycbz4Q0o4R3Kq9KubpgOSU5iy4eY6rcN2KcqGzo6GHi6hxZUM0bA/exec?"
            .add "results", "data"
            .add "treeSearch", True
            .add "ignore", vbNullString
            .add "timeout", 200
        End With
        
        With .add("tagsitejson")
            .add "restType", erSingleQuery
            .add "url", "https://googledrive.com/host/"
            .add "results", "data"
            .add "treeSearch", True
            .add "ignore", vbNullString
        End With

        With .add("colorschemer")
            .add "restType", erQueryPerRow
            .add "url", "https://script.google.com/macros/s/AKfycbzSdgK85uGHdQ9m076QkPV0B9a2kkgh7JHDmV8kzRgtkriSIwTn/exec?hex="
            .add "results", ""
            .add "treeSearch", True
            .add "ignore", vbNullString
            .add "timeout", 30

        End With
        With .add("fql")
            .add "restType", erSingleQuery
            .add "url", "http://graph.facebook.com/fql?q="
            .add "results", "data"
            .add "treeSearch", True
            .add "ignore", vbNullString
            .add "alwaysEncode", True
        End With
        With .add("fqlfeed")
            .add "restType", erSingleQuery
            .add "url", "https://graph.facebook.com/"
            .add "results", "data"
            .add "treeSearch", True
            .add "ignore", vbNullString
        End With

        With .add("villas")
            .add "restType", erSingleQuery
            .add "url", "http://www.villasofdistinction.com/tools/export-json/?destination="
            .add "results", vbNullString
            .add "treeSearch", True
            .add "ignore", vbNullString
            
        End With
        With .add("lukas")
            .add "restType", erSingleQuery
            .add "url", "http://somehere.com/someURL?someparameter="
            .add "results", vbNullString
            .add "treeSearch", True
            .add "ignore", vbNullString
        End With
        With .add("foobar")
            .add "restType", erSingleQuery
            .add "url", "http://somehere.com/someURL?someparameter="
            .add "results", "fooson"
            .add "treeSearch", True
            .add "ignore", vbNullString
        End With
        With .add("btc-e")
            .add "restType", erSingleQuery
            .add "url", "https://btc-e.com/api/2/ftc_btc/"
            .add "results", vbNullString
            .add "treeSearch", True
            .add "ignore", vbNullString
            
        End With
        With .add("btc-e-ticker")
            .add "restType", erSingleQuery
            .add "url", "https://btc-e.com/api/2/ftc_btc/"
            .add "results", "ticker"
            .add "treeSearch", True
            .add "ignore", vbNullString
        End With
        With .add("nestoria")
            .add "restType", erQueryPerRow
            .add "url", "http://api.nestoria.co.uk/api?country=uk&pretty=1&action=metadata&encoding=json&"
            .add "results", "response.metadata"
            .add "treeSearch", True
            .add "ignore", vbNullString
            
        End With
        With .add("publicstuff")
            .add "restType", erSingleQuery
            .add "url", "https://script.google.com/a/macros/mcpher.com/s/AKfycbzLXr1aQKQVK2imlIJp9C6m_HEBAmLBiYM28mfnLn_3oIe3c2kN/exec?entry="
            .add "results", "results"
            .add "treeSearch", False
            .add "ignore", vbNullString
        End With
        With .add("restserver")
            .add "restType", erSingleQuery
            .add "url", "?entry="
            .add "results", "restlibrary"
            .add "treeSearch", False
            .add "ignore", vbNullString
            .add "indirect", "publicstuff"
        End With
        With .add("duckduckgo")
            .add "restType", erSingleQuery
            .add "url", "http://api.duckduckgo.com/?format=json&q="
            .add "results", "relatedtopics"
            .add "treeSearch", True
            .add "ignore", vbNullString
        End With
        With .add("google patents")
            .add "restType", erSingleQuery
            .add "url", "https://ajax.googleapis.com/ajax/services/search/patent?v=1.0&rsz=8&q="
            .add "results", "responseData.results"
            .add "treeSearch", False
            .add "ignore", vbNullString
        End With
        With .add("twitter")
            .add "restType", erSingleQuery
            .add "url", "http://search.twitter.com/search.json?q="
            .add "results", "results"
            .add "treeSearch", False
            .add "ignore", vbNullString
        End With
        With .add("google books by isbn")
            .add "restType", erQueryPerRow
            .add "url", "https://www.googleapis.com/books/v1/volumes?q=isbn:"
            .add "results", "Items"
            .add "treeSearch", True
            .add "ignore", vbNullString
        End With
        With .add("rxNorm drugs")
            .add "restType", erSingleQuery
            .add "url", "http://rxnav.nlm.nih.gov/REST/drugs?name="
            .add "results", "drugGroup.conceptgroup.2.conceptProperties"
            .add "treeSearch", True
            .add "ignore", vbNullString
            .add "accept", "application/json"
        End With
        With .add("yahoo geocode")
            .add "restType", erQueryPerRow
            ' this was discontinued by yahoo
            '.add "url", "http://where.yahooapis.com/geocode?flags=J&location="
            .add "url", "http://gws2.maps.yahoo.com/findlocation?flags=J&location="
            .add "results", "ResultSet.Result"
            .add "treeSearch", True
            .add "ignore", vbNullString
        End With
        With .add("imdb by title")
            .add "restType", erQueryPerRow
            .add "url", "http://www.imdbapi.com/?tomatoes=true&t="
            .add "results", vbNullString
            .add "treeSearch", False
            .add "ignore", vbNullString
        End With
        With .add("itunes movie")
            .add "restType", erSingleQuery
            .add "url", "http://itunes.apple.com/search?entity=movie&media=movie&term="
            .add "results", "results"
            .add "treeSearch", False
            .add "ignore", vbNullString
        End With
        With .add("google finance")
            .add "restType", erQueryPerRow
            .add "url", "http://www.google.com/finance/info?infotype=infoquoteall&q="
            .add "results", "crest"
            .add "treeSearch", True
            .add "ignore", vbLf & "//"
        End With
        With .add("whatthetrend")
            .add "restType", erSingleQuery
            .add "url", "http://api.whatthetrend.com/api/v2/trends.json"
            .add "results", "trends"
            .add "treeSearch", False
            .add "ignore", vbNullString
        End With
        With .add("neildegrassetysonquotes")
            .add "restType", erSingleQuery
            .add "url", "http://www.neildegrassetysonquotes.com/quote_api/random"
            .add "results", ""
            .add "treeSearch", False
            .add "ignore", vbNullString
        End With
        With .add("tweetsentiments")
            .add "restType", erQueryPerRow
            .add "url", "http://data.tweetsentiments.com:8080/api/analyze.json?q="
            .add "results", "sentiment"
            .add "treeSearch", False
            .add "ignore", vbNullString
        End With
        With .add("topsy histogram")
            .add "restType", erQueryPerRow
            .add "url", "http://otter.topsy.com/searchhistogram.json?period=30&q="
            .add "results", "response"
            .add "treeSearch", True
            .add "ignore", vbNullString
        End With
        With .add("topsy count")
            .add "restType", erQueryPerRow
            .add "url", "http://otter.topsy.com/searchcount.json?q="
            .add "results", "response"
            .add "treeSearch", True
            .add "ignore", vbNullString
        End With
        With .add("tweetsentiment topics")
            .add "restType", erQueryPerRow
            .add "url", "http://data.tweetsentiments.com:8080/api/search.json?topic="
            .add "results", ""
            .add "treeSearch", True
            .add "ignore", vbNullString
        End With
        With .add("tweetsentiment details")
            .add "restType", erSingleQuery
            .add "url", "http://data.tweetsentiments.com:8080/api/search.json?topic="
            .add "results", "results"
            .add "treeSearch", True
            .add "ignore", vbNullString
        End With
        With .add("opencorporates reconcile")
            .add "restType", erSingleQuery
            .add "url", "http://opencorporates.com/reconcile?query="
            .add "results", "result"
            .add "treeSearch", True
            .add "ignore", vbNullString
        End With
        With .add("f1")
            .add "restType", erSingleQuery
            .add "url", "http://ergast.com/api/f1.json?limit="
            .add "results", "MRData.RaceTable.Races"
            .add "treeSearch", True
            .add "ignore", vbNullString
        End With
        With .add("f1 drivers")
            .add "restType", erSingleQuery
            .add "url", "http://ergast.com/api/f1/drivers.json?limit="
            .add "results", "MRData.DriverTable.Drivers"
            .add "treeSearch", True
            .add "ignore", vbNullString
        End With
        With .add("e-sim")
            .add "restType", erSingleQuery
            .add "url", "http://e-sim.org/apiMilitaryUnitMembers.html?id="
            .add "results", ""
            .add "treeSearch", True
            .add "ignore", vbNullString
        End With
        With .add("statwiki")
            .add "restType", erSingleQuery
            .add "url", "http://stats.grok.se/json/fr/"
            .add "results", "daily_views"
            .add "treeSearch", False
            .add "ignore", vbNullString
        End With
        With .add("GHStatListDB")
           .add "restType", erSingleQuery
           .add "url", "http://dl.dropbox.com/u/6341433/statlist.txt"
           .add "results", "result"
           .add "treeSearch", True
           .add "ignore", vbNullString
       End With

       With .add("craea")
            .add "restType", erRestType.erSingleQuery
            .add "url", "http://api.cscpro.org/esim/primera/tax/"
            .add "results", "tax"
            .add "treeSearch", True
            .add "ignore"
            .add "append", ".json"
        End With
        With .add("eSimResource")
            .add "restType", erRestType.erSingleQuery
            .add "url", "http://api.cscpro.org/esim/primera/market/"
            .add "results", "offer"
            .add "treeSearch", True
            .add "ignore"
            .add "append", ".json"
        End With
        With .add("battlenet")
            .add "restType", erRestType.erQueryPerRow
            .add "url", "http://us.battle.net/api/wow/item/"
            .add "results", ""
            .add "treeSearch", False
            .add "ignore"
        End With
        With .add("trello")
            .add "restType", erRestType.erSingleQuery
            .add "url", "https://api.trello.com/1/board/4ff1644acb179efe1718ec61?key=b5acff6f87bda62eba4ac7f6419fad20"
            .add "results", ""
            .add "treeSearch", True
            .add "ignore"
        End With
        With .add("huffingtonpost elections")
            .add "restType", erRestType.erSingleQuery
            .add "url", "http://elections.huffingtonpost.com/pollster/api/charts.json"
            .add "results", ""
            .add "treeSearch", True
            .add "ignore"
        End With
        With .add("jorum")
            .add "restType", erRestType.erSingleQuery
            .add "url", "http://dashboard.jorum.ac.uk/stats/"
            .add "results", ""
            .add "treeSearch", True
            .add "ignore"
        End With
        With .add("mercadolibre")
            .add "restType", erRestType.erSingleQuery
            .add "url", "https://api.mercadolibre.com/sites/MLA/search?q="
            .add "results", "results"
            .add "treeSearch", True
            .add "ignore"
        End With

        With .add("EC2")
            .add "restType", erRestType.erSingleQuery
            .add "url", "http://aws.amazon.com/ec2/pricing/pricing-reserved-instances.json"
            .add "results", "config.regions"
            .add "treeSearch", True
            .add "ignore", ""
        End With

        With .add("crunchbase relationships")
            .add "restType", erRestType.erSingleQuery
            .add "url", "http://api.crunchbase.com/v/1/person/"
            .add "results", "relationships"
            .add "treeSearch", True
            .add "ignore", ""
            .add "append", ".js"
        End With
        
        With .add("crunchbase companies")
            .add "restType", erRestType.erSingleQuery
            .add "url", "http://api.crunchbase.com/v/1/company/"
            .add "results", "relationships"
            .add "treeSearch", True
            .add "ignore", ""
            .add "append", ".js"
        End With
        With .add("scraperWiki")
            .add "restType", erRestType.erSingleQuery
            .add "url", "https://api.scraperwiki.com/api/1.0/scraper/search?format=jsondict&maxrows="
            .add "results", ""
            .add "treeSearch", False
            .add "ignore", ""
        End With
        With .add("who was In parliament")
            .add "restType", erRestType.erSingleQuery
            .add "url", "http://hansard.millbanksystems.com/all-members/"
            .add "results", ""
            .add "treeSearch", True
            .add "ignore", ""
        End With
        With .add("page rank")
            .add "restType", erRestType.erQueryPerRow
            .add "url", "http://prapi.net/pr.php?f=json&url="
            .add "results", ""
            .add "treeSearch", False
            .add "ignore", ""
        End With
        
        With .add("faa airport status")
            .add "restType", erRestType.erQueryPerRow
            .add "url", "http://services.faa.gov/airport/status/"
            .add "results", ""
            .add "treeSearch", False
            .add "ignore", ""
            .add "append", "?format=json"
        End With

        With .add("url shorten")
            .add "restType", erRestType.erQueryPerRow
            .add "url", "http://ttb.li/api/shorten?format=json&appname=ramblings&url="
            .add "results", ""
            .add "treeSearch", False
            .add "ignore", ""
            .add "append", ""
        End With
        
        With .add("uk postcodes")
            .add "restType", erRestType.erQueryPerRow
            .add "url", "http://www.uk-postcodes.com/postcode/"
            .add "results", ""
            .add "treeSearch", False
            .add "ignore", ""
            .add "append", ".json"
        End With
        With .add("freegeoip")
            .add "restType", erRestType.erQueryPerRow
            .add "url", "http://freegeoip.net/json/"
            .add "results", ""
            .add "treeSearch", False
            .add "ignore", ""
            .add "append", ""
        End With
        With .add("googlecurrencyconverter")
            .add "restType", erRestType.erQueryPerRow
            .add "url", "http://www.google.com/ig/calculator?hl=en&q=1USD=?"
            .add "results", ""
            .add "treeSearch", False
            .add "ignore", vbNullString
            .add "wire", True
        End With
        With .add("rate exchange")
            .add "restType", erRestType.erQueryPerRow
            .add "url", "http://rate-exchange.appspot.com/currency?from=USD&to="
            .add "results", ""
            .add "treeSearch", False
            .add "ignore", vbNullString
        End With

        
        With .add("scraperwikidata")
            .add "restType", erRestType.erSingleQuery
            .add "url", "https://api.scraperwiki.com/api/1.0/datastore/sqlite?format=jsondict&name="
            .add "results", ""
            .add "treeSearch", False
            .add "ignore", vbNullString
        End With
        With .add("urbarama")
            .add "restType", erRestType.erSingleQuery
            .add "url", "http://www.urbarama.com/api/project?sort=popular&offset=0&count=100&size=small&format=json"
            .add "results", "projects"
            .add "treeSearch", True
            .add "ignore", vbNullString
        End With
        With .add("urbaramamashup")
            .add "restType", erRestType.erSingleQuery
            .add "url", "?address="
            .add "results", "projects"
            .add "treeSearch", True
            .add "ignore", vbNullString
            .add "indirect", "publicstuff"
        End With
        With .add("builtwith")
            .add "restType", erRestType.erSingleQuery
            .add "url", "http://api.builtwith.com/api.json?lookup="
            .add "results", "Technologies"
            .add "treeSearch", False
            .add "ignore", vbNullString
            .add "append", "&key="
        End With
        With .add("ESRI Query")
            .add "restType", erQueryPerRow
            .add "url", "http://server.arcgisonline.com/ArcGIS/rest/services/Specialty/Soil_Survey_Map/MapServer/identify?geometryType=esriGeometryPoint&sr=4326&layers=1&time=&layerTimeOptions=&layerdefs=&tolerance=0&mapExtent=-119%2C38%2C-121%2C41&imageDisplay=400%2C300%2C96&returnGeometry=true&maxAllowableOffset=0&f=json&geometry="
            .add "results", "attributes"
            .add "treeSearch", True
            .add "ignore", vbNullString
        End With
        With .add("sina")
            .add "restType", erRestType.erSingleQuery
            .add "url", "http://stock.finance.sina.com.cn/usstock/api/json.php/US_MinKService.getMinK?type=15&___qn=3&symbol="
            .add "results", ""
            .add "treeSearch", True
            .add "ignore", vbNullString
        End With
        With .add("blister")
            .add "restType", erRestType.erSingleQuery
            .add "url", "https://script.google.com/a/macros/mcpher.com/s/AKfycbzhzIDmgY9BNeBu87puxMVUlMkJ4UkD_Yvjdt5MhOxR1R6RG88/exec?type=jsonp&source=scriptdb&module=blister&library="
            .add "results", "results"
            .add "treeSearch", True
            .add "ignore", vbNullString
            .add "append", "&query="
        End With
        With .add("blisterFunctions")
            .add "restType", erRestType.erSingleQuery
            .add "url", "https://script.google.com/macros/s/AKfycbzBskBK17poScDU9yHnfgmgPHyvgNejM3zxV7niGdhLeXPjw7Y4/exec"
            .add "results", ""
            .add "treeSearch", False
            .add "ignore", vbNullString
        End With
        With .add("postTest")
            .add "restType", erRestType.erSingleQuery
            .add "url", "http://posttestserver.com/post.php"
            .add "results", ""
            .add "treeSearch", False
            .add "ignore", vbNullString
        End With
        

    End With
    Set createRestLibrary = cj

End Function
Public Function restQuery(Optional sheetName As String = vbNullString, _
                    Optional sEntry As String = vbNullString, _
                    Optional sQuery As String = vbNullString, _
                    Optional sQueryColumn As String = vbNullString, _
                    Optional sRestUrl As String = vbNullString, _
                    Optional sResponseResults As String = vbNullString, _
                    Optional bTreeSearch As Boolean = True, _
                    Optional bPopulate As Boolean = True, _
                    Optional bClearMissing As Boolean = True, _
                    Optional complain As Boolean = True, _
                    Optional queryCanBeBlank As Boolean = False, _
                    Optional sFix As String = vbNullString, _
                    Optional user As String = vbNullString, _
                    Optional pass As String = vbNullString, _
                    Optional append As Boolean = False, _
                    Optional stampQuery As String = vbNullString, _
                    Optional appendQuery As String = vbNullString, _
                    Optional collectionNeeded As Boolean = True, _
                    Optional postData As String = vbNullString, _
                    Optional resultsFormat As erResultsFormat = erUnknown) As cRest
'   give it a known name, and somewhere to put the result
'   in the case where 1 query returns multiple rows, sQuery is the query contents
'   where 1 column contains the query for each row, sQueryColumn contains the name of the column
    Dim qType As erRestType, sUrl As String, sResults As String, sEntryType As erRestType, sc As ccell
    Dim dset As cDataSet, cr As cRest, sIgnore As String, cj As cJobject, cEntry As cJobject, job As cJobject
    Dim libAppend As String, _
        libAccept As String, bWire As Boolean, crIndirect As cRest, _
        rPlace As Range, bAlwaysEncode As Boolean, timeout As Long, oa As cOauth2
        
    libAppend = vbNullString
    libAccept = vbNullString
    
    Dim UA As cUAMeasure
    Set UA = registerUA("restQuery_" & sEntry)

    timeout = 0
    ' this is now a library object
    Set cEntry = getRestLibrary()
    
    If Not (sQuery = vbNullString Xor sQueryColumn = vbNullString) Then
        If Not queryCanBeBlank Then
            MsgBox ("you must provide one of either query contents or a query column name")
            Exit Function
        End If
    End If
    
    If Not (sEntry = vbNullString Xor sRestUrl = vbNullString) Then
        MsgBox ("you must provide one of either a known library entry or a rest URL")
        Exit Function
    End If

    ' based on whether a column name or a query argument was supplied
    If sQuery = vbNullString And Not queryCanBeBlank Then
        qType = erQueryPerRow
    Else
        qType = erSingleQuery
    End If
    ' get the characteristics from the crest library

    If sEntry = vbNullString Then
        sUrl = sRestUrl
        sResults = sResponseResults
        Set cj = New cJobject
        
    Else
        Set cj = cEntry.childExists(sEntry)
        If (cj Is Nothing) Then
            MsgBox (sEntry & " is not a known library entry")
            Exit Function
        End If
            
        sEntryType = cj.child("restType").toString
        sUrl = cj.child("url").toString
        sResults = cj.child("results").toString
        bTreeSearch = cj.child("treeSearch").toString = "True"
        sIgnore = cj.child("ignore").toString
        bAlwaysEncode = False
        If Not cj.childExists("timeout") Is Nothing Then timeout = cj.child("timeout").value
        If Not cj.childExists("alwaysEncode") Is Nothing Then bAlwaysEncode = cj.child("alwaysEncode").value
        If Not cj.childExists("append") Is Nothing Then libAppend = cj.child("append").toString
        If Not cj.childExists("accept") Is Nothing Then libAccept = cj.child("accept").toString
        If Not cj.childExists("wire") Is Nothing Then bWire = cj.child("wire").value
        If resultsFormat = erUnknown And _
            Not cj.childExists("resultsFormat") Is Nothing Then resultsFormat = cj.child("resultsFormat").value
        If Not cj.childExists("indirect") Is Nothing Then
            If cj.child("indirect").toString <> vbNullString Then
                ' now need to go off and execute that indirection - this could be recursive
                Set crIndirect = restQuery("", cj.child("indirect").toString, sEntry, , , , False)
                If crIndirect Is Nothing Then Exit Function
                sUrl = crIndirect.jObject.children("results").child("1.mystuff.publish").toString & sUrl
            End If
        End If
        If complain Then
            If abandonType(sEntry, qType, sEntryType) Then Exit Function
        End If
    End If
   
    If resultsFormat = erUnknown Then resultsFormat = erJSON
    
    Set cr = New cRest
    
    ' first we need to do oauth if its needed
    Set job = cj.childExists("authtype")
    If Not job Is Nothing And sFix = vbNullString Then
       
        If job.value = erOAUTH2 Then
            ' need to authorize and get token
            Set oa = getGoogled(cj.child("authScope").value)
            If (oa Is Nothing) Then Exit Function
        Else
            MsgBox ("Dont understand authtype " & CStr(job.value))
            Exit Function
        End If

    End If
    ' lets get the data
    Application.Calculation = xlCalculationManual
    If (sheetName <> vbNullString) Then
        Set dset = New cDataSet
        If (InStr(1, sheetName, "!") > 0) Then
            Set rPlace = Range(sheetName)
        Else
            Set rPlace = wholeSheet(sheetName)
        End If
        If (IsEmpty(rPlace.Cells(1, 1))) Then rPlace.Cells(1, 1).value = "crest"
        With dset.populateData(toEmptyBox(rPlace))
            ' ensure that the query column exists if it was asked for
            Dim sqa As Variant, si As Long, sqc As Collection
            If qType = erQueryPerRow Then
                Set sqc = New Collection
                sqa = Split(sQueryColumn, ",")
                For si = LBound(sqa) To UBound(sqa)
                    If Not .headingRow.validate(True, CStr(sqa(si))) Then Exit Function
                    sqc.add .headingRow.exists(CStr(sqa(si))), CStr(sqa(si))
                Next si
            End If
            If stampQuery <> vbNullString Then
                If Not .headingRow.validate(True, stampQuery) Then Exit Function
                Set sc = .headingRow.exists(stampQuery)
            End If
            ' alsmost there
            Set cr = cr.init(sResults, qType, sqc, _
                    , dset, bPopulate, sUrl, bClearMissing, _
                    bTreeSearch, complain, sIgnore, user, pass, append, sc, _
                    libAppend & appendQuery, libAccept, bWire, collectionNeeded, _
                    bAlwaysEncode, timeout, postData, resultsFormat, oa)
        End With

    Else
        Set cr = cr.init(sResults, qType, , _
                    , , False, sUrl, , _
                    bTreeSearch, complain, sIgnore, user, pass, append, sc, _
                    libAppend & appendQuery, libAccept, bWire, collectionNeeded, _
                    bAlwaysEncode, timeout, postData, resultsFormat, oa)
    End If
    

    If cr Is Nothing Then
        If complain Then MsgBox ("failed to initialize a rest class")
    Else
        Set cr = cr.execute(sQuery, sFix, complain)
        If cr Is Nothing Then
            If complain Then MsgBox ("failed to execute " & sQuery)
        Else
            Set restQuery = cr
        End If
    End If
    UA.postAppKill.tearDown

    Application.Calculation = xlCalculationAutomatic
End Function

Private Function abandonType(sEntry, qType As erRestType, targetType As erRestType) As Boolean

    If qType <> targetType Then
        abandonType = Not (vbYes = MsgBox(sEntry & " is normally " & _
                whichType(targetType) & _
                " but you have specified " & _
                whichType(qType) & ": try anyway?", vbYesNo))
    Else
        abandonType = False
    End If
End Function
Private Function whichType(t As erRestType) As String
    Select Case t
        Case erSingleQuery
            whichType = " single query that can return multiple rows"
        Case erQueryPerRow
            whichType = " a single column provides the query data for each row"
        Case Default
            Debug.Assert False
    End Select
End Function

Public Function createHeadingsFromKeys(job As cJobject, ds As cDataSet) As cDataSet
    ' use the keys of a cJobject as the headings
    Dim r As Range, jo As cJobject, dsNew As cDataSet
    ' clear the existing set
    Set r = ds.headingRow.where
    r.Worksheet.Cells.ClearContents
    If (Not job.hasChildren) Then
        MsgBox ("cjobject has no children to create headers from")
    Else
        For Each jo In job.children
            r.Offset(, jo.childIndex - 1).value = "'" & jo.key()
        Next jo
        Set dsNew = New cDataSet
        Set createHeadingsFromKeys = dsNew.populateData(r, , ds.name & "crest", , , , True)
    End If
            
End Function
Public Function getAndMakeJobjectFromXML(url As String) As cJobject
    ' we do an get on the given url
    Dim cb As cBrowser, helperUrl As String
    Set cb = New cBrowser
    helperUrl = _
     "https://script.google.com/macros/s/AKfycbziYOdWjNFtUR_TTQU-GiMYkan2h5ZDtaqeWIsYUAKEa6irjzNa/exec"

    With cb
        ' get the xml
        .httpGET url
        If .isOk Then
            Set getAndMakeJobjectFromXML = makeJobjectFromXML(.Text)
        Else
            MsgBox ("error getting " & url)
        End If
        .tearDown
    End With
    
End Function
Public Function makeJobjectFromXML(theXml As String, Optional complain As Boolean = True) As cJobject
    ' we do an get on the given url
    Dim cb As cBrowser, helperUrl As String
    Set cb = New cBrowser
    helperUrl = _
     "https://script.google.com/macros/s/AKfycbziYOdWjNFtUR_TTQU-GiMYkan2h5ZDtaqeWIsYUAKEa6irjzNa/exec"
   
    
    With cb

        'now convert to json using google apps script helper
        .httpPost helperUrl, theXml, True
        
        ' now we have it converted to json
        If .isOk Then
            With JSONParse(.Text)
                If .toString("status") = "good" Then
                    Set makeJobjectFromXML = .child("json")
                Else
                    If complain Then MsgBox (.toString("error") & " converting xml")
                End If
            End With
        Else
            If complain Then MsgBox (.status & " error getting xml convertor")
        End If
        

        .tearDown
    End With
End Function
Public Function getAndMakeJobjectAuto(url As String) As cJobject
    ' we do an get on the given url
    Dim cb As cBrowser, job As cJobject
    Set cb = New cBrowser
    
    With cb
        ' get the xml
        .httpGET url
        If .isOk Then
            ' try converting it from xml
            Set job = xmlStringToJobject(.Text, False)
            
            ' that didnt work, so assume its already json
            If job Is Nothing Then
                Set job = JSONParse(.Text)
            End If
            Set getAndMakeJobjectAuto = job
        
        Else
            MsgBox ("error getting " & url)
        End If
        .tearDown
    End With
    
End Function



