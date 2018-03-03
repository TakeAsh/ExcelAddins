# JoinRangeText

## JoinRangeText
Join texts in selected range.

## JoinRangeTextA
Join not empty texts in selected range.

## JoinRangeValue
Join values in selected range.

## JoinRangeValueA
Join not empty values in selected range.

## Sample
| A | B | C | D |
| ---- | ---- | ---- | ---- |
| =JoinRangeText($A$1:$C$3) | =JoinRangeTextA($A$1:$C$3) | =JoinRangeValue($A$1:$C$3) | =JoinRangeValueA($A$1:$C$3) |
| =JoinRangeText($A$1:$C$3, "$$") | =JoinRangeTextA($A$1:$C$3, "$$") | =JoinRangeValue($A$1:$C$3, "$$") | =JoinRangeValueA($A$1:$C$3, "$$") |
| =JoinRangeText($A$1:$C$3, CHAR(10)) | =JoinRangeTextA($A$1:$C$3, CHAR(10)) | =JoinRangeValue($A$1:$C$3, CHAR(10)) | =JoinRangeValueA($A$1:$C$3, CHAR(10)) |

![Sample](Sample.png)
(Align:Top, Text Wrap:On)
