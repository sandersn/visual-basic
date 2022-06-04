import keyboardinput.*;

package BaseConv;

public class BaseConv
{

	public BaseConv()	//NOT SURE YET IF I NEED THIS
	{
		super();
	}
	public static void main()	//THIS PROG TAKES NO ARGS. TOO BAD
	{
	keyboard in = new Keyboard();
	char conversion = '7';	//INIT TO 7 BY DEFAULT SINCE IT'S QUIT
	String answer = new String();
		while
		{
			System.out.println("Which conversion to do?\n"
				+ "1.Dec to Bin\n"
				+ "2.Bin to Dec\n"
				+ "3.Hex to Dec\n"
				+ "4.Dec to Hex\n"
				+ "5.Bin to Hex\n"
				+ "6.Hex to Bin\n"
				+ "7.Quit:\n");
				conversion = in.readString();
				if (conversion=='7')
				{
					System.out.println("Done.");
//					EXIT(1); //THIS IS C COMMAND
				}
				switch(conversion)
					case 1:
						answer = DecToBin(String.valueOf(conversion));	//NOT SURE THE VALUEOF
						//CONVERSION WORKS
					case 2:
					case 3:
					case 4:
					case 5:
					case 6:
					case 7:
			break;
		}
	}
}