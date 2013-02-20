import qualified Data.Map as M (fromList, lookup, findWithDefault)
import Data.List.Split (chunksOf, endBy)

import KpiStructure
import ExcelCom 

import Data.Aeson
import qualified Data.Text as T 
import Data.Aeson.Encode.Pretty
import qualified Data.ByteString.Lazy.Char8 as BL

    
fichierTest = "C:/Users/lyup7323/Developpement/Haskell/Com/qos1.xls"
fichierTest2 = "E:/Programmation/haskell/Com/qos1.xls"
fichierTest3 = "E:/Programmation/haskell/Com/qos.xls"
fichierTest4 = "C:/Users/lyup7323/Developpement/Haskell/Com/qos.xls"


{--
    - helpers
        - --} 

-- take 2 decimals

    -- reads double or put 0 
toDouble :: String -> Double
toDouble xs = case (reads.chgComma.endBy "," $ xs :: [(Double,String)] ) of
    [(d,s)] -> d
    _ -> 0
    where 
        chgComma [x,y] = x ++ "." ++ (take 2 y)
        chgComma xs = concat xs

toInt xs = case (reads xs :: [(Int,String)] ) of
    [(d,s)] -> d
    _ -> 0





{--
    - tests 
        - --}

sheetsName = ["BIV","BIC"]

main = coRun $ do
    (pExl, workBooks, workSheets) <- xlInit
    mapM_ (processRowData workSheets) sheetsName
    xlQuit workBooks pExl

xlQuit workBooks appXl = do
    workBooks # method_1_0 "Close" xlDoNotSaveChanges
    appXl # method_0_0 "Quit"

xlInit = do   
    pExl <- createObjExl
    workBooks <- pExl # getWorkbooks
    pExl # propertySet "DisplayAlerts" [inBool False]
    workBook <- workBooks # openWorkBooks fichierTest4
    putStrLn  $"File loaded: " ++ fichierTest4
    workSheets <- workBook # getWSheets'
    return (pExl, workBooks, workSheets)
    

processRowData :: Sheet a -> String -> IO ()
processRowData sheets sheetName = do 
    rowsService <- rowsFromSheet sheets sheetName
    putStrLn $ "got all datas from " ++ sheetName
    printListData rowsService
    putStrLn $ replicate 50 '-'


rowsFromSheet :: Sheet a -> String -> IO [String] 
rowsFromSheet workSheets sheetName= do 
    sheetSel <- workSheets # propertyGet_1 "Item" sheetName

    let row = 79
    let lastrow =  "C7:BC"++ show row
    putStrLn $ "endrow = " ++  lastrow
    rng <- sheetSel # propertyGet_1 "Range" lastrow
    fmap snd $ rng # enumVariants


printListData :: [String] -> IO ()    
printListData rows = do     
    let kpiData toX n  = map (toX.(rows!!).(+53*n)) [1..52]
        kpiName     = map (\n -> rows!!(53*n)) $ [0..72]
        kpiIndMap   = M.fromList $ zip kpiName [0..]
        rowNb s     = M.lookup s kpiIndMap

        rowSite     = rowNb "Sites"
        rowChannels = rowNb "Nb channels"
        rowMinutes  = rowNb "Minutes (Millions)"
        rowCalls    = rowNb "Calls (Millions)" 
        rowPgad     = rowNb "Post Gateway Answer Delay (sec)" 
        rowAsr      = rowNb "Answer Seizure Ratio (%)" 
        rowNer      = rowNb "Network Efficiency Ratio (%)" 
        rowAttps    = rowNb "ATTPS = Average Trouble Ticket Per Site" 
        rowAfis     = rowNb "Average FT Incident per Site\" AFIS" 
        rowMos      = rowNb "Mean Opinion Score (PESQ)" 
        rowPdd      = rowNb "Post Dialing Delay (sec)" 
        rowCsr      = rowNb "Call Sucessful Ratio" 
        rowRtd      = rowNb "RTD average" 
        rowAvail    = rowNb "Availability ratio HO (outage&changes)" 
        rowUnAvail  = rowNb "Unavailability minutes HO (outage&changes)" 
        rowComInd1  = rowNb "CommentIndispo1" 
        rowComInd2  = rowNb "CommentIndispo2" 
        rowComInd3  = rowNb "CommentIndispo3" 
        rowComInd4  = rowNb "CommentIndispo4" 
        rowComInd5  = rowNb "CommentIndispo5" 
        rowComAfis1 = rowNb "CommentAFIS1" 
        rowComAfis2 = rowNb "CommentAFIS2" 
        rowComMos1  = rowNb "CommentMOS1" 
        rowComMos2  = rowNb "CommentMOS2" 

        lookupData toX fill rowKpi = maybe (replicate 52 fill) -- default value [fill,...fill]
                                       (kpiData toX) -- handler 
                                       rowKpi -- Nothing or Just (kpi row) 

        nbSites = lookupData toInt 0 rowSite 
        nbChannels = lookupData toInt 0 rowChannels
        nbMinutes = lookupData toDouble 0 rowMinutes
        asr = lookupData toDouble 0 rowAsr
        ner = lookupData toDouble 0 rowNer
        attps = lookupData toDouble 0 rowAttps
        afis = lookupData toDouble 0 rowAfis
        mos = lookupData toDouble 0 rowMos
        pdd = lookupData toDouble 0 rowPdd
        csr = lookupData toDouble 0 rowCsr
        rtd = lookupData toDouble 0 rowRtd
        avail = lookupData toDouble 0 rowAvail
        unavail = lookupData toDouble 0 rowUnAvail
        
        commentIndisp1 = lookupData id "" rowComInd1
        commentIndisp2 = lookupData id ""  rowComInd2
        commentIndisp3 = lookupData id ""  rowComInd3
        commentIndisp4 = lookupData id ""  rowComInd4
        commentIndisp5 = lookupData id ""  rowComInd5
        commentAFIS1 = lookupData id ""  rowComAfis1
        commentAFIS2 = lookupData id ""  rowComAfis2
        commentMOS1 = lookupData id ""  rowComMos1
        commentMOS2 = lookupData id "" rowComMos2

    
    -- return (kpiName!! ind , kpiData ind)
    -- print them
    putStrLn "----nbSites---"
    print nbSites
    putStrLn "----nbChannels---"
    print nbChannels
    putStrLn "----nbMinutes---"
    print nbMinutes
    putStrLn "----ASR ---"
    print asr
    putStrLn "----Ner ---"
    print ner
    putStrLn "----mos ---"
    print mos
    putStrLn "----pdd ---"
    print pdd
    putStrLn "----csr ---"
    print csr
    putStrLn "---- rtd ---"
    print rtd
    putStrLn "---- commentIndisp1 ---"
    print commentIndisp1
    putStrLn "---- commentIndisp2 ---"
    print commentIndisp2
    putStrLn "---- commentIndisp3 ---"
    print commentIndisp3
    putStrLn "---- commentIndisp4 ---"
    print commentIndisp4
    putStrLn "---- commentIndisp5 ---"
    print commentIndisp5
    putStrLn "---- commentAFIS1 ---"
    print commentAFIS1
    putStrLn "---- commentAFIS2 ---"
    print commentAFIS2
    putStrLn "---- commentMOS1 ---"
    print commentMOS1
    putStrLn "---- CommentMOS2 ---"
    mapM_ putStr commentMOS2
    putStrLn "----attps ---"
    print attps
    putStrLn "----afis ---"
    print afis 
    putStrLn "----avail ---"
    print avail 
    putStrLn "----unvail ---"
    print unavail 
