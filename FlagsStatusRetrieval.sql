SELECT        Parks.UnitCode, Parks.Name, FieldNames.Label, FieldNames.Name AS Expr1, FieldNames.IsInSTATS, FieldNames.IsREC, FieldNames.IsRECH, FieldNames.IsNREC, FieldNames.IsNRECH, FieldNames.IsCL, 
                         FieldNames.IsCCG, FieldNames.IsBC, FieldNames.IsTT, FieldNames.IsTRVS, FieldNames.IsMISC, FieldNames.IsNROS, FieldNames.Formula
FROM            FieldNames INNER JOIN
                         Parks ON FieldNames.ParkId = Parks.ParkId