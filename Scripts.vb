'============================
'Adam's Auto Mapping Bullshit
'============================
'Notepad++ Remove Blank Lines/Rows Command
^\h*\R

'Labels
'14/3/F16(TDS)/16
'Query: DR_MH_ORD2

'Returns Primary
=VLOOKUP(A1,LocRef!$A$1:$C$999,2,TRUE)
'Returns Mailbox
=VLOOKUP(A1,LocRef!$A$1:$C$999,3,TRUE)

'Check if PC, if True return mailbox, else return Primary
'=IF (logical_test, value_if_true, value_if_false)
=IF(C1="PC",VLOOKUP(A1,LocRef!$A$1:$C$999,3,TRUE),VLOOKUP(A1,LocRef!$A$1:$C$999,2,TRUE))

'IF ELSE Logic
'=IF(T1,R1,IF(T2,R2,ELSE))
'Below line not used, failed logic test,keeping for reference
=IF(C1="CTN",VLOOKUP(A1,LocRef!$A$1:$C$999,2,TRUE),IF(C1="PC",VLOOKUP(A1,LocRef!$A$1:$C$999,3,TRUE),"Third"))


'Verbal Logic
'C1=U/M, J1
'A1=ItmNbr, I1

'
IF UOM = PC, CHECK NA STATUS, RETURN MB IF TRUE, RETURN PRIMARY IF FALSE, ELSE USE PRIMARY (ASSUME CTN)
'Formatted Verbal Logic

=IF(C1="PC",IF(("VLOOKUP(PIECE)")<>"na",("VLOOKUP(PIECE)"),"VLOOKUP(Primary)"),"VLOOKUP(Primary)")
'VB Logic Approximate

=IF(C1="PC",IF((VLOOKUP(A1,LocRef!$A$1:$C$999,3,TRUE))<>"na",(VLOOKUP(A1,LocRef!$A$1:$C$999,3,TRUE)),VLOOKUP(A1,LocRef!$A$1:$C$999,2,TRUE)),VLOOKUP(A1,LocRef!$A$1:$C$999,2,TRUE))
'VB Logic Exact (Test Script)

=IF(C1="PC",IF((VLOOKUP(A1,LocRef!$A$1:$C$999,3,FALSE))<>"na",(VLOOKUP(A1,LocRef!$A$1:$C$999,3,FALSE)),VLOOKUP(A1,LocRef!$A$1:$C$999,2,FALSE)),VLOOKUP(A1,LocRef!$A$1:$C$999,2,FALSE))
'VB Logic Exact (Crystal's Template)

=IF(J1="PC",IF((VLOOKUP(I1,LocRef!$A$1:$C$999,3,FALSE))<>"na",(VLOOKUP(I1,LocRef!$A$1:$C$999,3,FALSE)),VLOOKUP(I1,LocRef!$A$1:$C$999,2,FALSE)),VLOOKUP(I1,LocRef!$A$1:$C$999,2,FALSE))

'===
=IF(C1="PC",IF(VLOOKUP(A1,LocRef!$A$1:$C$999,3,TRUE)<>"na",VLOOKUP(A1,LocRef!$A$1:$C$999,3,TRUE),VLOOKUP(A1,LocRef!$A$1:$C$999,2,TRUE)),VLOOKUP(A1,LocRef!$A$1:$C$999,2,TRUE))
'===

'=======================
'Conditional Formatting
'=======================
'INDIRECT bypasses Excel error for referencing a different sheet // Query range must be in quotes
'If Logic
'=IF (T1, R_TRUE, R_FALSE)
'If AND Logic
'=IF(AND( T1 , T2 ), R_TRUE , R_FALSE )

'============================
'Quick Parts (Ref 2nd Column)
=MATCH($A1,INDIRECT("ItmClr!$A$2:$A$9999"),0)

'===============================
'Skirting/Vinyl Cartons (Ref 1st Column ItmClr)
=MATCH($A1,INDIRECT("ItmClr!$B$2:$B$9999"),0) 'default, does not test UoM of CTN
=IF(AND( $C1="CTN" , MATCH($A1,INDIRECT("ItmClr!$B$2:$B$9999"),0) ), 0 , 1 ) 'final, tests UoM/CTN

'===============================
'Skirting/Vinyl Pieces (Ref 1st Column ItmClr)
=MATCH($A1,INDIRECT("ItmClr!$B$2:$B$9999"),0) 'default, does not test UoM of PC
=IF(AND( $C1="PC" , MATCH($A1,INDIRECT("ItmClr!$B$2:$B$9999"),0) ), 0 , 1 ) 'final, tests UoM/PC

'===========================
'Wood Pieces (Ref 3rd Column)
=MATCH($A1,INDIRECT("ItmClr!$C$2:$C$9999"),0)