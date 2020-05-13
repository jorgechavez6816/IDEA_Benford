Sub Main
	Call BenfordsLaw()	'Ejemplo-Detalle de ventas.IMD
End Sub


' Análisis: Ley de Benford
Function BenfordsLaw
	FIRST1DIGIT_ANALYSIS = 1
	FIRST2DIGIT_ANALYSIS = 2
	FIRST3DIGIT_ANALYSIS = 4
	SECONDDIGIT_ANALYSIS = 8
	LAST2DIGIT_ANALYSIS = 16
	SECONDORDER_ANALYSIS = 32
	SUMMATION_ANALYSIS = 64
	POSITIVE_VALUES = 0
	NEGATIVE_VALUES = 1
	
	Set db = Client.OpenDatabase("Ejemplo-Detalle de ventas.IMD")
	Set task = db.BenfordsLaw
	task.FieldToUse = "TOTAL"
	task.ValueType = POSITIVE_VALUES
	task.CheckBoundaries = TRUE
	task.MADConclusion = TRUE
	dbName = "Benford_primer_digito.IMD"
	task.AddAnalysis FIRST1DIGIT_ANALYSIS, dbname
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase(dbName)
End Function