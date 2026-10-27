
VB Declaration Converter   v 1.2
Copyright Exile Software 1995

Public Domain
Please notify me of incorrent or unusable declarations which you may discover!
Thanks.

1. Description

    The Declaration Converter  reads WIN32 SDK header files and converts C prototypes to VB declarations.
    The Converter also reads "win32api.txt" and similar text files containing VB declarations and sorts and converts them to a form suitable for importing into a database.
     The VB declarations are separated into lists of Constants, Types, Functions, and Subs and sorted alphabetically.  The declarations can then be processed to shorten then.  The declaration lists can be saved to a single file, or to four seperate files.

2. Usage

    Reading VB files:    The Converter can read MS or other texts of VB declarations.    Declarations are extracted and sorted and can be processed to shorten them.    Keywords are stripped ("Type/End Type", "Declare Sub", "Declare Function", "Const") so that the output can be used with a tool which looks up API declarations.    (This speeds searches, reduces file sizes, and provides flexibility when pasting the declarations into VB projects.)

    Reading C headers:    The Converter can read WIN32 SDK header files and convert C prototypes to VB declarations.  See the notes below before using these declarations.

    Processing:    The declarations can be processed in the following ways:
       	1. Comments can be stripped
        	2. Datatypes can be converted to the shortened form (e.g.  "As Long" > "&"). (Functions and Subs only)
        	3. Variable names can be shortened. (Functions and Subs only) 
    
    Editing:    Individual items can be edited, added, removed, or copied to the clipboard.

************************************************************************************************************************************

Notes on converting C prototypes to VB declarations


1. Untranslated terms.   In several cases, C expressions are not translated.  This may happen if there is no VB equivalent of the expression, or if a decision is required by the user before the declaration can be used.
        a. Bit-fields are not translated. (See 'alphaops.h')
        b. Unions are not translated.  You must remove the union declaration and replace it with one or more variables which fit your intended usage.
        c. Structures and unions which occur within structures are not translated.  Again, replace these with one or variables of the same size.
        d. Unrecognized C datatypes are not translated.   These datatypes will be declared by their original name.

2. Unknowns.   Untranslatable lines are added to an "Unknown" list.  Check this list whenever you translate a file.

3. 'Lib'.   Many of the 'Lib xxx' statements may be inaccurate.  Because there is no indication in the header files of the dll which contains the exported functions, in most cases the converter will write:
        Lib "x.dll"
    Some headers contain comments indicating the dll to which the applies; in these cases, the Converter writes a dll name.   However, these names are not necessarily accurate.   The dll's may not even exist or may not be present on your machine.
   The next version will correctly supply the 'Lib' name.   It will include a master list of exported functions and the ability to dynamically verify from which dll a given function is exported.   For now you must use another tool to produce a list of exported functions from the WIN32 dlls, then look up the function in this list to find the correct dll.

4. Arrays.   C arrays which are sized by C constants will be 1 too large. This is compensated for by declaring arrays " ( 1 to x)".

5. Pointers in structures.   C Structures which contain pointers to structures should not be used from VB. Ex: "HD_NOTIFY", from "commctrl.h".   C Structures which contain pointers to strings should not be used. Ex: "TBSAVEPARAMSA", from "commctrl.h".  Structures containing pointers are marked with a warning comment when translated, and are translated "As Long". 
    Note:   You can use these structures as VB Types, but you must supply the address of these variables at runtime.   (A simple way to do this is to pass the variable to a dll and have it return the address.)

6. ANSI/Wide functions.  Functions which have both an ANSI and a Wide form are aliased to the ANSI version.   Wide versions of functions are not translated.

7. ANSI/Wide structures.   Both ANSI and Wide structures are translated.   The Wide version of a structure should not ordinarily be used from VB.

8. LARGE_INTEGER.   64-bit values are translated as LARGE_INTEGER structures.

9. 'TCHAR'.   Datatypes which can be compiled to ANSI or Wide are interpreted as ANSI (one-byte).   Character types are translated as Byte.   You may want to change some of these to fixed-length Strings.

10. Inline functions.   The prototype for inline functions is translated.   The function body is ignored.

11. Constants.   Most "#define" statements are translated as constants.   Macros are ignored.   The 'Unknown' list will contain all the statements which were not tranlsated.

12. Enumerated types.   Enumerated types are not translated.

13 OLE.   A lot of OLE-related stuff is not translated because it can't be used from VB.   Again, the 'Unknown' list will contain these statements.   Some Interface declarations and structures will be not be coherent.   These should just be ignored, as they can't be used from VB anyway.

************************************************************************************************************************************

Function Parameters.

1. 'As Any'.    'As Any' is used by MS to flag functions which contain parameters which can accept a string or numerical argument.  You should never use this in VB; instead create aliases such as FooByNum/FooByString. 
    I have found no way to detect such parameters; LPVOID is translated 'As Any'.   You must create the aliases when you find you need them.

2. 'ByRef/ByVal'.   This version supplies the 'ByRef' or 'ByVal' qualifier for function declarations.   C Integers and Longs are translated as Long.   Character types are translated as Byte or String.

3. Pointers to Arrays.   Pointers to arrays are translated as the array's datatype and passed ByRef.   Pass the first element of the array when calling the function.   Ex: "ArrayName(0)".

4. Pointers to numbers.   In general, pointers are translated as the datatype they point to and passed ByRef.   This causes VB to create a pointer and pass it to the function.

5. Pointers to Strings.   Strings work exactly the opposite to numbers.  Pointers to strings are translated as String passed ByVal.   VB will create a pointer to the String and pass it to the function.

6. Pointers to functions.   Pointers to functions are translated as Long.   You must have a dll containing the callback function to use these functions from VB, and the dll will probably manage such calls itself.

7. Variable names.   Most C function prototypes list only the datatype of parameters.   Arbitrary names are provided for variables (x1,x2,...).     

8. Unrecognized datatypes.   If a datatype is not recognized, it is transliterated.   Check the 'Unknown' list or the header file if the datatype is not apparent.  

9. Hardcoded argument values.    A few functions have hardcoded argument values.   These are rendered with the C value as the datatype.  Ex: ( ' x as 0x16 ' )

************************************************************************************************************************************

Known Problems.


1. Some constructions cannot be parsed correctly. Consider the following, from "winnt.h":
	#if defined(MIDL_PASS)
	typedef struct _LARGE_INTEGER {
	#else // MIDL_PASS
	typedef union _LARGE_INTEGER {
   	 struct {
       	 DWORD LowPart;
       	 LONG HighPart;
  	  };
    	struct {
       	 DWORD LowPart;
       	 LONG HighPart;
   	 } u;
	#endif //MIDL_PASS
	    LONGLONG QuadPart;
	} LARGE_INTEGER;
The converter returns this unchanged, but only because it recognizes it explicitly.   Other such constructions will cause parsing errors, because the converter performs text parsing.  It ignores preprocessing directives.   It works against the WIN32 SDK headers because it was designed for the style of constructions which appear therein, but keep this in mind if translating other header files.

2. In general, the converter requires you to manually check for accuracy and usabilty before attempting to incorpoate the output in a VB project.   I haven't systematically checked the output for usability, so check the SDK documentation when a problem arises. 

3. I am trying to bring the output to as close a match to correct VB declarations as can be gotten from an examination of header files.   If you have any suggestions, please post me a message.

Ed@scling.hk.super.net
CS: 100267,546
