Attribute VB_Name = "Module4"
Public Function InterpolateAndGetSUMIN(fromTime As Variant, toTime As Variant, TimeTable As Variant, IntegralTable As Variant)
    'InterpolateAndGetSUMIN = Worksheets("���Ȫ�").Evaluate("=(INDEX(Table11[�ɶ�],MATCH(($B$2+1),Table11[�ɶ�],1))*INDEX(Table11[Integral],MATCH(($B$2+1),Table11[�ɶ�],1))+INDEX(Table11[�ɶ�],MATCH(($B$2+1),Table11[�ɶ�],1)+1)*INDEX(Table11[Integral],MATCH(($B$2+1),Table11[�ɶ�],1)+1))/((INDEX(Table11[�ɶ�],MATCH(($B$2+1),Table11[�ɶ�],1))+INDEX(Table11[�ɶ�],MATCH(($B$2+1),Table11[�ɶ�],1)+1)))-(INDEX(Table11[�ɶ�],MATCH($B$2,Table11[�ɶ�],1))*INDEX(Table11[Integral],MATCH($B$2,Table11[�ɶ�],1))+INDEX(Table11[�ɶ�],MATCH($B$2,Table11[�ɶ�],1)+1)*INDEX(Table11[Integral],MATCH($B$2,Table11[�ɶ�],1)+1))/((INDEX(Table11[�ɶ�],MATCH($B$2,Table11[�ɶ�],1))+INDEX(Table11[�ɶ�],MATCH($B$2,Table11[�ɶ�],1)+1)))")

    InterpolateAndGetSUMIN = Application.Evaluate("")
End Function


