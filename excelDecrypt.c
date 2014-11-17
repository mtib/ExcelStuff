#include <stdlib.h>
#include <stdio.h>

#if defined _WIN16 || defined _WIN32 || defined _WIN64
#define EOL "\r\n"
#elif defined __linux__ || defined __unix__
#define EOL "\n"
#endif

static int decryptFile(const char*, const char*);

int main(int argc,char **argv)
{
	if(argc>2)
	{
		switch(decryptFile(argv[1],argv[2]))
		{
			case -1:
				printf("Error opening File %s" EOL,argv[1]);
				break;
			case -2:
				printf("Error opening File %s" EOL,argv[2]);
				break;
			default:
				puts("Password removed\r\nDo the folowing now:" EOL
				"1. Open the excel sheet and press [ALT]+[F11]" EOL
				"   Confirm any errors that might appear" EOL
				"2. go to Tools > VBA project properties." EOL
				"3. In the Tab \"Protection\" enter any password" EOL
				"   Do not clear the Checkbox!" EOL
				"4. Save, close the Editor, save the excel sheet and close it" EOL
				"5. Open it again" EOL
				"6. Repear step 1 and 2, there should be no errors" EOL
				"7. Clear the password checkbox in the \"protection\" tab" EOL
				"8. Repeat Steps 4 and 5." EOL
				"9. Your password is gone!" EOL EOL
				"/u/AyrA_ch");
				break;
		}
	}
	else
	{
		puts("excelDecrypt <input-filename> <output-filename>");
	}
	return 0;
}

int decryptFile(const char* fName1,const char* fName2)
{
	char* cc;
	FILE* fp;
	int fs,c,i;
	if(fp=fopen(fName1,"rb"))
	{
		fseek(fp, 0L, SEEK_END);
		cc=malloc(fs=ftell(fp));
		fseek(fp, 0L, SEEK_SET);
		fread(cc,sizeof(cc[0]),fs,fp);
		for(i=0;i<fs-4;i++)
		{
			if(cc[i]=='D' && cc[i+1]=='P' && cc[i+2]=='B' && cc[i+3]=='=' && cc[i+4]=='"')
			{
				cc[i+2]='x';
			}
		}
		fclose(fp);
		
		if(fp=fopen(fName2,"wb"))
		{
			fwrite(cc,sizeof(cc[0]),fs,fp);
			fclose(fp);
		}
		else
		{
			free(cc);
			return -2;
		}
		free(cc);
	}
	else
	{
		return -1;
	}
	return 0;
}
