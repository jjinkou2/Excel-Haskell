import Control.Concurrent

t = replicate 10 "Toto"

main = do 
    box <- newMVar 0
    sequence $ map (forkIO. thPrint box) t  
    boucle box

thPrint box str = do 
    putStrLn $ "trhead: " ++ str
    val <- takeMVar box
    putMVar box $ val + 1

boucle box = do 
    val <- takeMVar box 
    if val >= length t 
    then do
        putStrLn "finished"
        return () 
    else do 
        putMVar box val
        print "waiting"
        boucle box

