{-# LANGUAGE OverloadedStrings, RecordWildCards #-}

import Control.Applicative ((<$>), (<*>), empty)
import Data.Aeson
import qualified Data.Text as T 
import Data.Aeson.Encode.Pretty
import qualified Data.ByteString.Lazy.Char8 as BL

data Coord = Coord { x :: Double, y :: Double }
             deriving (Show)

data KpiVal = KpiNum 

-- A ToJSON instance allows us to encode a value as JSON.

instance ToJSON Coord where
  toJSON (Coord xV yV) = object [ "x" .= xV,
                                  "y" .= yV ]

-- A FromJSON instance allows us to decode a value from JSON.  This
-- should match the format used by the ToJSON instance.

mos = decode  "{\"toto\": [4.5,2.3,4.9] }"::Maybe Value
mos2 = toJSON ([4.3,2.3,4.9]::[Double])
comment= toJSON  (["Bon","Mauvais","Tres Bon"]::[String])
kpiMos = KpiStruct "mos" mos2
kpiCommentMos = KpiStruct "CommentMos" comment

data KpiStruct = KpiStruct {kpiStructName :: T.Text, kpiStructData :: Value} deriving (Show)

instance ToJSON KpiStruct where
    toJSON KpiStruct{..} = object [kpiStructName .= kpiStructData]
--test = BL.putStrLn.encodePretty.toJSON $ ["toto","toti"]

instance FromJSON Coord where
  parseJSON (Object v) = Coord <$>
                         v .: "x" <*>
                         v .: "y"
  parseJSON _          = empty

test = BL.putStrLn.encode $ kpiCommentMos

main :: IO ()
main = do
  let req = decode "{\"x\":3.0,\"y\":-1.0}" :: Maybe Coord
  print req
  let reply = Coord 123.4 20
  BL.putStrLn (encode reply)
