{-# LANGUAGE BangPatterns, OverloadedStrings #-}
import Data.ByteString.Lex.Double (readDouble)
 
import Data.Vector.Binary
import Data.Binary
import qualified Data.ByteString.Char8 as L
import qualified Data.Vector.Unboxed as V
 
main = do
    s <- L.readFile "dat"
    let v = parse s :: V.Vector Int
    print  v
 
-- Fill a new vector from a file containing a list of numbers.
parse = V.unfoldr step
  where
     step !s = case L.readInt s of
        Nothing       -> Nothing
        Just (!k, !t) -> Just (k, L.tail t)



toDouble :: L.ByteString -> Double
toInt xs = case L.readInt xs of
    Just (d,s) -> d
    Nothing -> 0

toDouble xs = case (readDouble.L.intercalate ".".L.split ',' $ xs)  of
    Just (d,s) -> d
    Nothing -> 0

{-
    let kpiData toX n  = map (toX.(rows!!).(+53*n)) [1..52]
        kpiName     = map (\n -> rows!!(53*n)) $ [0..72]
        kpiIndMap   = M.fromList $ zip kpiName [0..]
-- take 2 decimals
trunc :: Double -> Double
trunc double = (fromInteger $ round $ double * (10^2)) / (10.0 ^^2)

    -- reads double or put 0 
toDouble :: String -> Double
toDouble xs = case (reads.chgComma.endBy "," $ xs :: [(Double,String)] ) of
    [(d,s)] -> d
    _ -> 0
    where 
        chgComma [x,y] = x ++ "." ++ (take 2 y)
        chgComma xs = concat xs
        -}
