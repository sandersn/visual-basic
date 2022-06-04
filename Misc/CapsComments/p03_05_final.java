/************************
PROBLEM DEFINITION:
~~~~~~
Input:empName, hoursWorked, hourRate, payrollDeductionPercent
Output:empName, hoursWorked, hourRate, grossPay, payrollDeduction, netPay
Processing:
	Figure grossPay from hoursWorked times hourRate
	Figure payrollDeduction from payrollDeductionPercent times grossPay
	Figure netPay by removing payrollDeduction from grossPay
~~~~~~
ALGORITHMS:
~~~~~~
main()
BEGIN
	readEmployeeInformation()
	calcGrossPay()
	calcPayrollDeduction()
	calcNetPay()
	writeEmployeeInformation()
END
readEmployeeInformation()
BEGIN
	WRITE program description message
	WRITE userprompt
	READ empName
	WRITE userprompt
	READ hoursWorked
	WRITE userprompt
	READ hourRate
	WRITE userprompt
	READ payrollDeductionPercent
END
calcGrossPay()
BEGIN
	grossPay = hoursWorked * hourRate
END
calcPayrollDeduction()

BEGIN
	payrollDeduction = payrollDeductionPercent * grossPay
END
calcNetPay()
BEGIN
	netPay = grossPay - payrollDeduction
END
writeEmployeeInformation()
BEGIN
	WRITE userprompt
	WRITE empName
	WRITE userprompt
	WRITE hoursWorked
	WRITE userprompt
	WRITE hourRate
	WRITE userprompt
	WRITE grossPay
	WRITE userprompt
	WRITE payrollDeduction
	WRITE userprompt
	WRITE netPay
END
~~~~~~
*************************/
public class p03_05
{
}//END CLASS P03_05
//THE END
//THE END TEST TWO
//NOTE:I HAVE A GIANT SCREEN. PLEASE FORGIVE SUPER LONG LINES, KTHX
//or maybe I will just print this out in Times New Roman or something