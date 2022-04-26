# region IMPORTS
from __future__ import annotations
import datetime
import time
from numbers import Number
from typing import Any, Optional, List, Tuple, TypeVar, overload, Union
from typing_extensions import Literal
from com.sun.star.awt import XWindow
from com.sun.star.awt import XControl
from com.sun.star.awt import XControlModel
from com.sun.star.awt.tree import XMutableTreeNode
from com.sun.star.awt.tree import XMutableTreeDataModel
from com.sun.star.beans import PropertyValue
from com.sun.star.chart import XDiagram
from com.sun.star.document import XEmbeddedScripts
from com.sun.star.drawing import XShape
from com.sun.star.form import XForm
from com.sun.star.frame import XDesktop
from com.sun.star.lang import XComponent
from com.sun.star.script import XLibraryContainer
from com.sun.star.script.provider import XScriptProvider
from com.sun.star.sheet import XSheetCellCursor
from com.sun.star.sheet import XSpreadsheet
from com.sun.star.sdb import DatabaseDocument
from com.sun.star.sdbc import XConnection as UNOXConnection
from com.sun.star.sdbc import XDatabaseMetaData
from com.sun.star.table import XCellRange
from com.sun.star.table import XTableChart
from com.sun.star.uno import XInterface
from com.sun.star.uno import XComponentContext
from com.sun.star.util import DateTime as UNODateTime
from com.sun.star.util import Date as UNODate
from com.sun.star.util import Time as UNOTime
from com.sun.star.form import ListSourceType
# endregion IMPORTS

# region Types
_T = TypeVar("_T")
# endregion Types

# region MetaClass
class _Singleton(type):
    """
    A Singleton metaclass design pattern
    Credits: « Python in a Nutshell » by Alex Martelli, O'Reilly
    """

    instances: dict
    def __call__(cls: _T, *args, **kwargs) -> _T: ...
# endregion MetaClass

# region ScriptForge Class
class ScriptForge(object, metaclass=_Singleton):
    """
        The ScriptForge (singleton) class encapsulates the core of the ScriptForge run-time
            - Bridge with the LibreOffice process
            - Implementation of the inter-language protocol with the Basic libraries
            - Identification of the available services interfaces
            - Dispatching of services
            - Coexistence with UNO

        It embeds the Service class that manages the protocol with Basic
    """

    # region Class attributes
    hostname: str
    port: int
    componentcontext: XComponentContext
    scriptprovider: XScriptProvider
    SCRIPTFORGEINITDONE: bool
    # endregion Class attributes

    # region Class constants
    library: Literal["ScriptForge"]
    Version: Literal["7.3"]
    # endregion Class constants
    # Basic dispatcher for Python scripts
    basicdispatcher: Literal[
        "@application#ScriptForge.SF_PythonHelper._PythonDispatcher"
    ]
    # Python helper functions module
    pythonhelpermodule: Literal["ScriptForgeHelper.py"]
    #
    # region VarType() constants
    V_EMPTY: Literal[0]
    V_NULL: Literal[1]
    V_INTEGER: Literal[2]
    V_LONG: Literal[3]
    V_SINGLE: Literal[4]
    V_DOUBLE: Literal[5]
    V_CURRENCY: Literal[6]
    V_DATE: Literal[7]
    V_STRING: Literal[8]
    V_OBJECT: Literal[9]
    V_BOOLEAN: Literal[11]
    V_VARIANT: Literal[12]
    V_ARRAY: Literal[8192]
    V_ERROR: Literal[-1]
    V_UNO: Literal[16]
    # Object types
    objMODULE: Literal[1]
    objCLASS: Literal[2]
    objUNO: Literal[3]
    # endregion VarType() constants
    # region Special argument symbols
    cstSymEmpty: Literal["+++EMPTY+++"]
    cstSymNull: Literal["+++NULL+++"]
    cstSymMissing: Literal["+++MISSING+++"]
    # endregion Special argument symbols
    # Predefined references for services implemented as standard Basic modules
    servicesmodules: dict
    def __init__(self, hostname: str = ..., port: int = ...):
        """
        Because singleton, constructor is executed only once while Python active
        Arguments are mandatory when Python and LibreOffice run in separate processes

        Args:
            hostname (str, optional): probably 'localhost'. Defaults to ''.
            port (int, optional): Port Number. Defaults to 0.
        """
    # region ClassMethods
    @classmethod
    def ConnectToLOProcess(
        cls, hostname: str = ..., port: int = ...
    ) -> XComponentContext:
        """
        Called by the ScriptForge class constructor to establish the connection with
        the requested LibreOffice instance
        The default arguments are for the usual interactive mode

        Args:
            hostname (Optional[str], optional):  probably 'localhost'. Defaults to ''.
            port (Optional[int], optional): Port number. Defaults to 0.
        """
    @classmethod
    def ScriptProvider(cls, context: XComponentContext | None = ...) -> XScriptProvider:
        """
        Returns the general script provider

        Args:
            context (XComponentContext): component context.

        Returns:
            XScriptProvider: the general script provider
        """
    @classmethod
    def InvokeSimpleScript(cls, script: str, *args) -> Any:
        """
        Create a UNO object corresponding with the given Python or Basic script
        The execution is done with the invoke() method applied on the created object
        Implicit scope: Either::

            "application"   a shared library                (BASIC)
            "share"         a library of LibreOffice Macros (PYTHON)

        Args:
            script (str): See Note

        Note:
            Arg ``script`` is Eithor:

            - [@][scope#][library.]module.method - Must not be a class module or method
            - [@] means that the targeted method accepts ParamArray arguments (Basic only)
            - [scope#][directory/]module.py$method - Must be a method defined at module level

        Returns:
            Any:the value returned by the invoked script, or an error if the script was not found
        """
    @classmethod
    def InvokeBasicService(
        cls, basicobject: object, flags: int, method: str, *args: Any
    ) -> Union[tuple, datetime.datetime]:
        """
        Execute a given Basic script and interpret its result
        This method has as counterpart the ScriptForge.SF_PythonHelper._PythonDispatcher() Basic method

        Args:
            basicobject (object):  a Service subclass
            flags (int): see the vb* and flg* constants in the SFServices class
            method (str): the name of the method or property to invoke, as a string

        Other Args:
            the arguments of the method. Symbolic cst* constants may be necessary
        
        Returns:
            Union[tuple, list]: The invoked Basic counterpart script (with InvokeSimpleScript()) will return a tuple

        Note:

            Return info:

                [0]     The returned value - scalar, object reference or a tuple

                [1]     The Basic VarType() of the returned value
                        Null, Empty and Nothing have different vartypes but return all None to Python

                Additionally, when [0] is a tuple:
                    [2]     Number of dimensions in Basic

                Additionally, when [0] is a UNO or Basic object:
                    [2]     Module (1), Class instance (2) or UNO (3)

                    [3]     The object's ObjectType

                    [4]     The object's ServiceName

                    [5]     The object's name
            
            When an error occurs Python receives None as a scalar. This determines the occurrence of a failure
            The method returns either:
                - the 0th element of the tuple when scalar, tuple or UNO object
                - a new Service() object or one of its subclasses otherwise
        """
    # endregion ClassMethods
    # region Static Methods
    @staticmethod
    def SetAttributeSynonyms() -> None:
        """
        A synonym of an attribute is either the lowercase or the camelCase form of its original ProperCase name.
        In every subclass of SFServices:

        1) Fill the propertysynonyms dictionary with the synonyms of the properties listed in serviceproperties
            Example:
                .. code::

                    serviceproperties = dict(ConfigFolder = False, InstallFolder = False)
                    propertysynonyms = dict(configfolder = 'ConfigFolder', installfolder = 'InstallFolder',
                                            configFolder = 'ConfigFolder', installFolder = 'InstallFolder')

        2) Define new method attributes synonyms of the original methods
            Example:
                .. code::

                def CopyFile(...):
                    # etc ...
                copyFile, copyfile = CopyFile, CopyFile
        """
    # endregion Static Methods
# endregion ScriptForge Class

# region SFServices CLASS    (ScriptForge services superclass)
class SFServices:
    """
    Generic implementation of a parent Service class
    Every service must subclass this class to be recognized as a valid service
    A service instance is created by the CreateScriptService method
    It can have a mirror in the Basic world or be totally defined in Python

    Every subclass must initialize 3 class properties:
        servicename (e.g. 'ScriptForge.FileSystem', 'ScriptForge.Basic')
        servicesynonyms (e.g. 'FileSystem', 'Basic')
        serviceimplementation: either 'python' or 'basic'
    This is sufficient to register the service in the Python world

    The communication with Basic is managed by 2 ScriptForge() methods:
        InvokeSimpleScript(): low level invocation of a Basic script. This script must be located
            in a usual Basic module. The result is passed as-is
        InvokeBasicService(): the result comes back encapsulated with additional info
            The result is interpreted in the method
            The invoked script can be a property or a method of a Basic class or usual module
    It is up to every service method to determine which method to use

    For Basic services only:
        Each instance is identified by its
            - object reference: the real Basic object embedded as a UNO wrapper object
            - object type ('SF_String', 'DICTIONARY', ...)
            - class module: 1 for usual modules, 2 for class modules
            - name (form, control, ... name) - may be blank

        The role of the SFServices() superclass is mainly to propose a generic properties management
        Properties are got and set following next strategy:
            1. Property names are controlled strictly ('Value' or 'value', not 'VALUE')
            2. Getting a property value for the first time is always done via a Basic call
            3. Next occurrences are fetched from the Python dictionary of the instance if the property
                is read-only, otherwise via a Basic call
            4. Read-only properties may be modified or deleted exceptionally by the class
                when self.internal == True. The latter must immediately be reset after use

        Each subclass must define its interface with the user scripts:
        1.  The properties
            Property names are proper-cased
            Conventionally, camel-cased and lower-cased synonyms are supported where relevant
                a dictionary named 'serviceproperties' with keys = (proper-cased) property names and value = boolean
                    True = editable, False = read-only
                a list named 'localProperties' reserved to properties for internal use
                    e.g. oDlg.Controls() is a method that uses '_Controls' to hold the list of available controls
            When:
                    forceGetProperty = False    # Standard behaviour
            read-only serviceproperties are buffered in Python after their 1st get request to Basic
            Otherwise set it to True to force a recomputation at each property getter invocation
            If there is a need to handle a specific property in a specific manner:
                @property
                def myProperty(self):
                    return self.GetProperty('myProperty')
        2.   The methods
            a usual def: statement
                def myMethod(self, arg1, arg2 = ''):
                    return self.Execute(self.vbMethod, 'myMethod', arg1, arg2)
            Method names are proper-cased, arguments are lower-cased
            Conventionally, camel-cased and lower-cased homonyms are supported where relevant
            All arguments must be present and initialized before the call to Basic, if any
    """
    # region CONST
    vbGet: Literal[2]
    vbLet: Literal[4]
    vbMethod: Literal[1]
    vbSet: Literal[8]
    flgPost: Literal[32]
    """The method or the property implies a hardcoded post-processing"""
    flgDateArg: Literal[64]
    """Invoked service method may contain a date argument"""
    flgDateRet: Literal[128]
    """Invoked service method can return a date"""
    flgArrayArg: Literal[512]
    """1st argument can be a 2D array"""
    flgArrayRet: Literal[1024]
    """Invoked service method can return a 2D array (standard modules) or any array (class modules)"""
    flgUno: Literal[256]
    """Invoked service method/property can return a UNO object"""
    flgObject: Literal[2048]
    """1st argument may be a Basic object"""
    # Basic class type
    moduleClass: Literal[2]
    moduleStandard: Literal[1]
    # endregion CONST
    
    # region Attribs
    forceGetProperty: bool = False
    """Define the default behaviour for read-only properties: buffer their values in Python"""
    propertysynonyms: dict = ...
    """Empty dictionary for lower/camelcased homonyms or properties"""
    internal_attributes: tuple
    # Shortcuts to script provider interfaces
    SIMPLEEXEC: Any
    EXEC: Any
    # endregion Attribs

    # region Methods
    def __init__(
        self,
        reference: int = ...,
        objtype: str | None = ...,
        classmodule: int = ...,
        name: str = ...,
    ) -> None:
        """
         Trivial initialization of internal properties
        If the subclass has its own __init()__ method, a call to this one should be its first statement.
        Afterwards localProperties should be filled with the list of its own properties

        Args:
            reference (int, optional):the index in the Python storage where the Basic object is stored. Defaults to -1.
            objtype (str, optional): ('SF_String', 'DICTIONARY', ...). Defaults to None.
            classmodule (int, optional):  Module (1), Class instance (2). Defaults to 0.
            name (str, optional): '' when no name. Defaults to ''.
        """
    def __getattr__(self, name: str) -> Any | None: ...
    def __setattr__(self, name: str, value: Any) -> None: ...
    def __repr__(self) -> str: ...
    def Dispose(self) -> None:
        """
        Dispose
        """
    def ExecMethod(self, flags: int = ..., methodname: str = ..., *args: Any) -> Any: ...
    def GetProperty(self, propertyname: str, arg=None):
        """
        Get the given property from the Basic world
        """
    def Properties(self) -> list: ...
    # def basicmethods(self) -> Any: ...
    # def basicproperties(self) -> Any: ...
    def SetProperty(self, propertyname: str, value: Any) -> Any:
        """
        Set the given property to a new value in the Basic world
        """
    # endregion Methods
# endregion SFServices CLASS    (ScriptForge services superclass)

# region SFScriptForge CLASS    (alias of ScriptForge Basic library)
class SFScriptForge:
    ...
    # region SF_Array CLASS
    class SF_Array(SFServices, metaclass=_Singleton):
        """
        Provides a collection of methods for manipulating and transforming arrays of one dimension (vectors)
        and arrays of two dimensions (matrices). This includes set operations, sorting,
        importing to and exporting from text files.
        The Python version of the service provides a single method: ``ImportFromCSVFile()``

        Note:
            Because Python has built-in list and tuple support, most of the methods in the Array
            service are available for Basic scripts only. The only exception is ``ImportFromCSVFile()``
            which is supported in both Basic and Python.

        See Also:
            `Array Help <https://tinyurl.com/y8b7bq2d>`_
        """
        def ImportFromCSVFile(
            self, filename: str, delimiter: str = ..., dateformat: str = ...
        ) -> Any:
            """
            Difference with the Basic version: dates are returned in their iso format,
            not as any of the datetime objects.

            Args:
                filename (str): The name of the text file containing the data.
                    The name must be expressed according to the current FileNaming property of the
                    SF_FileSystem service.
                delimiter (str, optional): A single character, usually, a comma, a semicolon or a TAB character
                    (Default = ",").
                dateformat (str, optional): A special mechanism handles dates when dateformat is either
                    "YYYY-MM-DD", "DD-MM-YYYY" or "MM-DD-YYYY". The dash (-) may be replaced by a dot (.),
                    a slash (/) or a space. Other date formats will be ignored. Dates defaulting to an empty
                    string "" are considered as normal text.

            See Also:
                `SF_Array Help ImportFromCSVFile <https://tinyurl.com/yuc67r32#ImportFromCSVFile>`_
            """
    # endregion SF_Array CLASS

    # region SF_Basic CLASS
    class SF_Basic(SFServices, metaclass=_Singleton):
        """
        This service proposes a collection of Basic methods to be executed in a Python context
        simulating the exact syntax and behaviour of the identical Basic builtin method.
        Typical example:
            SF_Basic.MsgBox('This has to be displayed in a message box')

        The signatures of Basic builtin functions are derived from
            core/basic/source/runtime/stdobj.cxx

        See Also:
           `ScriptForge.Basic Service <https://tinyurl.com/ycv7q52r>`_
        """
        # region CONST
        # Basic helper functions invocation
        module: Literal["SF_PythonHelper"]
        # Message box constants
        MB_ABORTRETRYIGNORE: Literal[2]
        MB_DEFBUTTON1: Literal[128]
        MB_DEFBUTTON2: Literal[258]
        MB_DEFBUTTON3: Literal[215]
        MB_ICONEXCLAMATION: Literal[48]
        MB_ICONINFORMATION: Literal[64]
        MB_ICONQUESTION: Literal[32]
        MB_ICONSTOP: Literal[16]
        MB_OK: Literal[0]
        MB_OKCANCEL: Literal[1]
        MB_RETRYCANCEL: Literal[5]
        MB_YESNO: Literal[4]
        MB_YESNOCANCEL: Literal[3]
        IDABORT: Literal[3]
        IDCANCEL: Literal[2]
        IDIGNORE: Literal[5]
        IDNO: Literal[7]
        IDOK: Literal[1]
        IDRETRY: Literal[4]
        IDYES: Literal[6]
        # endregion CONST

        # region Methods
        @classmethod
        def CDate(cls, datevalue: Any) -> Union[datetime.datetime, object]:
            """
            Converts a numeric expression or a string to a datetime.datetime Python native object.

            datevalue: (Any): a numeric expression or a string representing a date.
            Note:
                This method exposes the Basic builtin function `Basic CDate <https://tinyurl.com/2p8sc4sy>`_ to Python scripts.

            See Also:
                `SF_Basic Help CDate <https://tinyurl.com/ycv7q52r#CDate>`_
            """
        @staticmethod
        def CDateFromUnoDateTime(
            unodate: Union[UNODateTime, UNODate, UNOTime],
        ) -> Union[datetime.datetime, object]:
            """
            Converts a UNO date/time representation to a datetime.datetime Python native object
            
            Args:
                unodate (UNODateTime, UNODate, UNOTime): uno date object.

            Returns:
                Union[datetime, object]: the equivalent datetime.datetime

            Note:
                Arg ``unodate`` can be one of the following:
                    - com.sun.star.util.DateTime
                    - com.sun.star.util.Date
                    - com.sun.star.util.Time
            See Also:
                `SF_Basic Help CDateFromUnoDateTime <https://tinyurl.com/ycv7q52r#CDateFromUnoDateTime>`_
            """
        @staticmethod
        def CDateToUnoDateTime(
            date: Union[float, time.struct_time, datetime.datetime, datetime.date, datetime.time],
        ) -> Union[UNODateTime, Any]:
            """
            Converts a date representation into the ccom.sun.star.util.DateTime date format

            Args:
                date (float |time.localtime | datetime | date | time ]): datetime like object

            Returns:
                UNODateTime: a com.sun.star.util.DateTime

            Note:
                When arg ``date`` is a ``float`` it is considered a ``time.time`` value.

            See Also:
                `SF_Basic Help CDateToUnoDateTime <https://tinyurl.com/ycv7q52r#CDateToUnoDateTime>`_
            """
        @classmethod
        def ConvertFromUrl(cls, url: str) -> str:
            """
            Convert from url

            Args:
                url (str): a string representing a file in URL format

            Returns:
                str: The same file name in native operating system notation
            
            Example:
                .. code::
                
                    a = bas.ConvertFromUrl('file:////boot.sys')

            See Also:
                `SF_Basic Help ConvertFromUrl <https://tinyurl.com/ycv7q52r#ConvertFromUrl>`_
            """
        @classmethod
        def ConvertToUrl(cls, systempath: str) -> str:
            """
            Convert to url

            Args:
                systempath (str): a string representing a file in native operating system notation

            Returns:
                str: The same file name in URL format
            
            Exampe:
                .. code::
                
                    >>> a = bas.ConvertToUrl('C:\\boot.sys')
                    >>> print(a)
                    'file:///C:/boot.sys'

            See Also:
                `SF_Basic Help ConvertToUrl <https://tinyurl.com/ycv7q52r#ConvertToUrl>`_
            """
        @classmethod
        def CreateUnoService(cls, servicename: str) -> XInterface:
            """
            Creates a uno service

            Args:
                servicename (str):  a string representing the service to creat

            Returns:
                XInterface: A UNO object
            
            Example:
                .. code::
                
                    a = bas.CreateUnoService('com.sun.star.i18n.CharacterClassification')

            See Also:
                `SF_Basic Help CreateUnoService <https://tinyurl.com/ycv7q52r#CreateUnoService>`_
            """
        @classmethod
        def DateAdd(
            cls,
            interval: str,
            number: int,
            date: Union[float, time.struct_time, datetime.datetime, datetime.date, datetime.time],
        ) -> Union[UNODateTime, Any]:
            """
            Adds a date or time interval to a given date/time a number of times and returns the resulting date.

            Args:
                interval (str): A string expression from the following table, specifying the date or time interval.
                number (int): A numerical expression specifying how often the interval value will be added when
                    positive or subtracted when negative.
                date (Union[float, time.struct_time, datetime.datetime, datetime.date, datetime.time]): A given
                    datetime.datetime value, the interval value will be added number times to this datetime.datetime value.
            Returns:
                datetime.datetime: A datetime.datetime value.

            Note:
                Arg ``interval`` valid expression values:
                    - yyyy - year
                    - q - Quarter
                    - m - Month
                    - y - Day of year
                    - w - Weekday
                    - ww - Week of year
                    - d - Day
                    - h - Hour
                    - n - Minute
                    - s - Second

            See Also:
                `SF_Basic Help DateAdd <https://tinyurl.com/ycv7q52r#DateAdd>`_
            """
        @classmethod
        def DateDiff(
            cls,
            interval: str,
            date1: Union[float, time.struct_time, datetime.datetime, datetime.date, datetime.time],
            date2: Union[float, time.struct_time, datetime.datetime, datetime.date, datetime.time],
            firstdayofweek: int = ...,
            firstweekofyear: int = ...,
        ) -> int:
            """
            Gets the number of date or time intervals between two given date/time values.

            Args:
                interval (str): A string expression specifying the date interval, as detailed in above DateAdd method.
                date1 (float | time.struct_time | datetime.datetime | datetime.date | datetime.time): The first datetime.datetime values to be compared.
                date2 (float | time.struct_time | datetime.datetime | datetime.date | datetime.time): The second datetime.datetime values to be compared.
                firstdayofweek (int, optional): An optional parameter that specifies the starting day of a week.
                firstweekofyear (int, optional): An optional parameter that specifies the starting week of a year.

            Note:
                Arg ``interval`` valid expression values:
                    - yyyy - year
                    - q - Quarter
                    - m - Month
                    - y - Day of year
                    - w - Weekday
                    - ww - Week of year
                    - d - Day
                    - h - Hour
                    - n - Minute
                    - s - Second

                Arg ``firstdayofweek`` values:
                    - 0 - Use system default value
                    - 1 - Sunday (default)
                    - 2 - Monday
                    - 3 - Tuesday
                    - 4 - Wednesday
                    - 5 - Thursday
                    - 6 - Friday
                    - 7 - Saturday

                Arg ``firstweekofyear`` values:
                    - 0 - Use system default value
                    - 1 - Week 1 is the week with January, 1st (default)
                    - 2 - Week 1 is the first week containing four or more days of that year
                    - 3 - Week 1 is the first week containing only days of the new year

            See Also:
                `SF_Basic Help DateDiff <https://tinyurl.com/ycv7q52r#DateDiff>`_

           Returns:
               int: A Number
           """
        @classmethod
        def DatePart(
            cls,
            interval: str,
            date: Union[time.struct_time, datetime.datetime, datetime.date, datetime.time, str],
            firstdayofweek: int = ...,
            firstweekofyear: int = ...,
        ) -> int:
            """
            Gets a specified part of a date.

            Args:
                interval (str): The unit of the date interval. See Note.
                date (Union[ time.struct_time, datetime.datetime, datetime.date, datetime.time, str ]): The date from which to extract a part
                firstdayofweek (int, optional): the starting day of a week. Defaults to 1.
                firstweekofyear (int, optional): the starting week of a year. Defaults to 1.

            Returns:
                int: The specified part of the date
            
            Example:
                .. code-block:: python
                
                    >>> a = bas.DatePart('y', bas.Now()) # day of year, 2022
                    >>> print(a)
                    98

            Note:
                Arg ``interval`` valid expression values:
                    - yyyy - year
                    - q - Quarter
                    - m - Month
                    - y - Day of year
                    - w - Weekday
                    - ww - Week of year
                    - d - Day
                    - h - Hour
                    - n - Minute
                    - s - Second

                Arg ``firstdayofweek`` values:
                    - 0 - Use system default value
                    - 1 - Sunday (default)
                    - 2 - Monday
                    - 3 - Tuesday
                    - 4 - Wednesday
                    - 5 - Thursday
                    - 6 - Friday
                    - 7 - Saturday

                Arg ``firstweekofyear`` values:
                    - 0 - Use system default value
                    - 1 - Week 1 is the week with January, 1st (default)
                    - 2 - Week 1 is the first week containing four or more days of that year
                    - 3 - Week 1 is the first week containing only days of the new year

            See Also:
                `SF_Basic Help DatePart <https://tinyurl.com/ycv7q52r#DatePart>`_
            """
        @classmethod
        def DateValue(cls, string: str) -> Union[datetime.datetime, Any]:
            """
            Computes a date value from a date string.

            Args:
                string (str): String expression that contains the date that you want to calculate.

            Returns:
                Union[datetime.datetime, object]: The converted date
            
            Example:
                .. code-block:: python
                
                    >>> a = bas.DateValue('2021-02-18')
                    >>> print(a)
                    datetime.datetime(2021, 2, 18, 0, 0)

            See Also:
                `DateValue <https://tinyurl.com/ycv7q52r#DateValue>`_
            """
        @classmethod
        def Format(
            cls, expression: Union[datetime.datetime, Number], format: str = ...
        ) -> str:
            """
            Converts a number to a string, and then formats it according to the format that you specify.

            Args:
                expression (datetime.datetime | Number): Numeric expression that you want to convert to a formatted string.
                format (str, optional): the format to apply. Defaults to "".

            Returns:
                str: The formatted value
            
            Example:
                .. code-block:: python
                
                    >>> a =  bas.Format(6328.2, '##,##0.00')
                    >>> print(a)
                    '6,328.20'

            See Also:
                `SF_Basic Help Format <https://tinyurl.com/ycv7q52r#Format>`_
            """
        @classmethod
        def GetDefaultContext(cls) -> XComponentContext:
            """
            Gets the default context of the process service factory, if existent, else returns a null reference.

            See Also:
                `SF_Basic Help GetDefaultContext <https://tinyurl.com/ycv7q52r#GetDefaultContext>`_
            """

        @classmethod
        def GetGuiType(cls) -> int:
            """
            Gets a numerical value that specifies the graphical user interface.

            This function is only provided for backward compatibility with previous versions.

            Returns:
                int: The GetGuiType value, 1 for Windows, 4 for UNIX
            
            Example:
                .. code-block:: python
                
                    >>> print(bas.GetGuiType())
                    1

            See Also:
                `SF_Basic Help GetGuiType <https://tinyurl.com/ycv7q52r#GetGuiType>`_
            """
        @classmethod
        def GetPathSeparator(cls) -> str:
            """
            Gets the operating system-dependent directory separator used to specify file paths.

            Use os.pathsep from os Python module to identify the path separator.

            Returns:
                str: os path separator

            See Also:
                `SF_Basic Help GetPathSeparator <https://tinyurl.com/ycv7q52r#GetPathSeparator>`_
            """
        @classmethod
        def GetSystemTicks(cls) -> int:
            """
            Gets the number of system ticks provided by the operating system.

            You can use this function to optimize certain processes.
            Use this method to estimate time in milliseconds:

            Returns:
                int: system ticks

            See Also:
                `SF_Basic Help GetSystemTicks <https://tinyurl.com/ycv7q52r#GetSystemTicks>`_
            """
        class GlobalScope(metaclass=_Singleton):
            @classmethod  # Mandatory because the GlobalScope class is normally not instantiated
            def BasicLibraries(cls) -> XLibraryContainer:
                """
                Gets the UNO object containing all shared Basic libraries and modules.
                This method is the Python equivalent to GlobalScope.BasicLibraries in Basic scripts.

                Returns:
                    XLibraryContainer: Uno object.

                See Also:
                `SF_Basic Help BasicLibraries <https://tinyurl.com/ycv7q52r#BasicLibraries>`_
                """
            @classmethod
            def DialogLibraries(cls) -> XLibraryContainer:
                """
                Gets the UNO object containing all shared dialog libraries.

                Returns:
                    DialogLibraryContainer: Uno object

                See Also:
                `SF_Basic Help DialogLibraries <https://tinyurl.com/ycv7q52r#DialogLibraries>`_
                """
        @classmethod
        def InputBox(
            cls,
            prompt: str,
            title: str = ...,
            default: str = ...,
            xpostwips: int = ...,
            ypostwips: int = ...,
        ) -> str:
            """
            Displays an input box.

            Args:
                prompt (str): String expression displayed as the message in the dialog box.
                title (str): String expression displayed in the title bar of the dialog box.
                default (str): String expression displayed in the text box as default if
                    no other input is given.
                xpostwips (int): Integer expression that specifies the horizontal position of the dialog.
                    The position is an absolute coordinate and does not refer to the window of LibreOffice.
                ypostwips: Integer expression that specifies the vertical position of the dialog.
                    The position is an absolute coordinate and does not refer to the window of LibreOffice.

            Note:
                If ``xpostwips`` and ``ypostwips`` are omitted, the dialog is centered on the screen.
                The position is specified in `twips <https://tinyurl.com/j3bueabr#twips>`_.

            Returns:
                str: string.

            See Also:
                `SF_Basic Help InputBox <https://tinyurl.com/ycv7q52r#InputBox>`_
            """
        @classmethod
        def MsgBox(cls, prompt: str, buttons: int = ..., title: str = ...) -> int:
            """
            Displays a dialogue box containing a message and returns an optional value.
            
            MB_xx constants help specify the dialog type, the number and type of buttons to display,
            plus the icon type. By adding their respective values they form bit patterns, that define the
            ``MsgBox`` dialog appearance.

            Args:
                prompt (str): String expression displayed as a message in the dialog box.
                buttons (int, optional):Any integer expression that specifies the dialog type, as well as the number and type of buttons to display, and the icon type. buttons represents a combination of bit patterns, that is, a combination of elements can be defined by adding their respective values. Defaults to 0.
                title (str, optional): String expression displayed in the title bar of the dialog. Defaults to "".

            Note:
                Arg ``buttons`` constants
                ::
                    Named constant          Integer value   Definition
                    MB_OK                   0               Display OK button only.
                    MB_OKCANCEL             1               Display OK and Cancel buttons.
                    MB_ABORTRETRYIGNORE     2               Display Abort, Retry, and Ignore buttons.
                    MB_YESNOCANCEL          3               Display Yes, No, and Cancel buttons.
                    MB_YESNO                4               Display Yes and No buttons.
                    MB_RETRYCANCEL          5               Display Retry and Cancel buttons.
                    MB_ICONSTOP             16              Add the Stop icon to the dialog.
                    MB_ICONQUESTION         32              Add the Question icon to the dialog.
                    MB_ICONEXCLAMATION      48              Add the Exclamation Point icon to the dialog.
                    MB_ICONINFORMATION      64              Add the Information icon to the dialog.
                    MB_DEFBUTTON1           128             First button in the dialog as default button.
                    MB_DEFBUTTON2           256             Second button in the dialog as default button.
                    MB_DEFBUTTON3           512             Third button in the dialog as default button.

                **Return Value**
                ::
                    Named constant   Integer value   Definition
                    IDOK             1              OK
                    IDCANCEL         2              Cancel
                    IDABORT          3              Abort
                    IDRETRY          4              Retry
                    IDIGNORE         5              Ignore
                    IDYES            6              Yes
                    IDNO             7              No

            Returns:
                int: The pressed button as int.
            Example:
                .. code-block:: python
                
                    >>> a = bas.MsgBox ('Please press a button:', bas.MB_ICONEXCLAMATION, 'Dear User')
                    >>> print(a)
                    1

            See Also:
                `SF_Basic Help MsgBox <https://tinyurl.com/ycv7q52r#MsgBox>`_
            """
        @classmethod
        def Now(cls) -> datetime.datetime:
            """
            Gets the current system date and time as a ``datetime.datetime`` Python native object.

            See Also:
                `SF_Basic Help Now <https://tinyurl.com/ycv7q52r#Now>`_
            """
        @classmethod
        def RGB(cls, red: int, green: int, blue: int) -> int:
            """
            Gets an integer color value consisting of red, green, and blue components.

            Args:
                red (int): Any integer expression that represents the red component (0-255) of the composite color.
                green (int): Any integer expression that represents the green component (0-255) of the composite color.
                blue (int): Any integer expression that represents the blue component (0-255) of the composite color.

            Returns:
                int: an integer color value consisting of red, green, and blue components.

            See Also:
                `SF_Basic Help RGB <https://tinyurl.com/ycv7q52r#RGB>`_
            """

        @overload
        @classmethod
        def Xray(cls) -> None: ...
        @overload
        @classmethod
        def Xray(cls, unoobject: object) -> None:
            """
            Inspect Uno objects or variables.

            Args:
                unoobject (object):A variable or UNO object.

            See Also:
                `SF_Basic Help Xray <https://tinyurl.com/ycv7q52r#Xray>`_
            """
        # endregion Methods

        # region Properties
        @property
        def StarDesktop(self) -> XDesktop: ...
        starDesktop, stardesktop = StarDesktop, StarDesktop
        @property
        def ThisComponent(self) -> Union[XComponent, None]:
            """
            If the current component refers to a LibreOffice document, this method
            returns the UNO object representing the document.
            
            Returns:
                XComponent | None: the current component or None when not a document

            See Also:
                `SF_Basic Help ThisComponent <https://tinyurl.com/ycv7q52r#ThisComponent>`_
            """
        thisComponent, thiscomponent = ThisComponent, ThisComponent
        @property
        def ThisDatabaseDocument(self) -> Union[XEmbeddedScripts, None]:
            """
            If the script is being executed from a Base document or any of its subcomponents
            this method returns the main component of the Base instance.
            
            Returns:
                 XEmbeddedScripts | None: the current Base (main) component or
                    None when not a Base document or one of its subcomponents

            See Also:
                `SF_Basic Help ThisDatabaseDocument <https://tinyurl.com/ycv7q52r#ThisDatabaseDocument>`_
            """
        thisDatabaseDocument, thisdatabasedocument = (
            ThisDatabaseDocument,
            ThisDatabaseDocument,
        )
        # endregion Properties
    # endregion SF_Basic CLASS
    
    # region SF_Dictionary CLASS
    class SF_Dictionary(SFServices, dict):
        """
            The service adds to a Python dict instance the interfaces for conversion to and from
            a list of UNO PropertyValues

            Usage:
                dico = dict(A = 1, B = 2, C = 3)
                myDict = CreateScriptService('Dictionary', dico)    # Initialize myDict with the content of dico
                myDict['D'] = 4
                print(myDict)   # {'A': 1, 'B': 2, 'C': 3, 'D': 4}
                propval = myDict.ConvertToPropertyValues()
            or
                dico = dict(A = 1, B = 2, C = 3)
                myDict = CreateScriptService('Dictionary')          # Initialize myDict as an empty dict object
                myDict.update(dico) # Load the values of dico into myDict
                myDict['D'] = 4
                print(myDict)   # {'A': 1, 'B': 2, 'C': 3, 'D': 4}
                propval = myDict.ConvertToPropertyValues()
            
            See Also: 
                `ScriptForge.Dictionary service <https://tinyurl.com/y9quuboc>`_
            """

        def __init__(self, dic: Optional[dict] = None) -> None: ...
        def ConvertToPropertyValues(self) -> List[PropertyValue]:
            """
            Store the content of the dictionary in an array of PropertyValues.
            Each entry in the list is a ``com.sun.star.beans.PropertyValue``.
            the key is stored in Name, the ``value`` is stored in ``Value``.

            If one of the items has a type ``datetime``, it is converted to a ``com.sun.star.util.DateTime``
            structure. If one of the items is an empty list, it is converted to ``None``.

            The resulting list is empty when the dictionary is empty.

            Returns:
                List[PropertyValue]: List of property values.

            See Also:
                `SF_Dictionary Help ConvertToPropertyValues <https://tinyurl.com/y9quuboc#ConvertToPropertyValues>`_
            """
        def ImportFromPropertyValues(
            self, propertyvalues: Tuple[PropertyValue, ...], overwrite: bool = False
        ) -> bool:
            """
            Inserts the contents of an array of ``PropertyValue`` objects into the current dictionary.
            ``PropertyValue`` Names are used as keys in the dictionary, whereas Values contain
            the corresponding values. Date-type values are converted to ``datetime.datetime`` instances.

            Args:
                propertyvalues (Tuple[PropertyValue, ...]): tuple containing com.sun.star.beans.PropertyValue objects
                overwrite (bool, optional): When True, entries with same name may exist in the dictionary and their values
                    are overwritten. When False (default), repeated keys are not overwritten. Defaults to False.

            Returns:
                bool: True when successful

            See Also:
                `SF_Dictionary Help ImportFromPropertyValues <https://tinyurl.com/y9quuboc#ImportFromPropertyValues>`_
            """
    # endregion SF_Dictionary CLASS
    
    # region SF_Exception CLASS
    class SF_Exception(SFServices, metaclass=_Singleton):
        """
        The Exception service is a collection of methods for code debugging and error handling.

        The Exception service console stores events, variable values and information about errors.
        Use the console when the Python shell is not available, for example in Calc user defined functions (UDF)
        or during events processing.
        Use DebugPrint() method to aggregate additional user data of any type.

        Console entries can be dumped to a text file or visualized in a dialogue.

        See Also:
            `ScriptForge.Exception service <https://tinyurl.com/y8ezar7q>`_
        """
        # region Methods
        def Console(self, modal: bool = ...) -> Any: ...
        def ConsoleClear(self, keep: int = ...) -> Any: ...
        def ConsoleToFile(self, filename: str) -> Any: ...
        def DebugDisplay(self, *args) -> Any: ...
        def DebugPrint(self, *args) -> Any: ...
        @classmethod
        def PythonShell(cls, variables: dict | None = ...) -> None:
            """
            Opens an APSO Python shell as a non-modal window.
            The Python script keeps running after the shell is opened.
            The output from ``print`` statements inside the script are shown in the shell.

            Only a single instance of the APSO Python shell can be opened at any time.
            Hence, if a Python shell is already open, then calling this method will have no effect.

            Args:
                variables (dict, None): a Python dictionary with variable names and values that will be
                    passed on to the APSO Python shell. By default, all local variables are passed using
                    Python's builtin locals() function.

            See Also:
                `SF_Exception Help PythonShell <https://tinyurl.com/y8ezar7q#PythonShell>`_
            """
        @classmethod
        def RaiseFatal(cls, errorcode: str, *args: Any) -> None:
            """
            Generate a run-time error caused by an anomaly in a user script detected by ScriptForge
            The message is logged in the console. The execution is STOPPED
            
            For INTERNAL USE only
            """
        # endregion Methods

        # region Properties
        @property
        def Description(self) -> str:
            """
            Gets/Sets the error message text.

            Default value is '' or a string containing the Basic run-time error message.
            """
        @property
        def Number(self) -> int:
            """
            Gets/Sets the location in the code where the error occurred. It can be a numeric value or text.

            Default value is 0 or the numeric value corresponding to the Basic run-time error code.
            """
        @property
        def Source(self) -> int:
            """
            Gets/Sets the location in the code where the error occurred. It can be a numeric value or text.

            Default value is ``0`` or the code line number for a standard Basic run-time error.
            """
        # endregion Properties
    # endregion SF_Exception CLASS
    
    # region SF_FileSystem CLASS
    class SF_FileSystem(SFServices, metaclass=_Singleton):
        """
        The "FileSystem" service includes common file and folder handling routines.
        
        See Also:
            `ScriptForge.FileSystem service <https://tinyurl.com/ybxpt7eo>`_
        """
          # region Methods
        def BuildPath(self, foldername: str, name: str) -> str:
            """
            Joins a folder path and the name of a file and returns the full file name with a
            valid path separator. The path separator is added only if necessary.

            Args:
                foldername (str): The path with which name will be combined.
                    The specified path does not need to be an existing folder.
                name (str): The name of the file to be appended to foldername. This parameter uses
                    the notation of the current operating system.

            Returns:
                str: The path concatenated with the file name after insertion of a path separator, if necessary

            See Also:
                `SF_FileSystem Help BuildPath <https://tinyurl.com/ybxpt7eo#BuildPath>`_
            """
        def CompareFiles(
            self, filename1: str, filename2: str, comparecontents: bool = ...
        ) -> bool:
            """
            Compare 2 files and return ``True`` if they seem identical.

            Depending on the value of the ``comparecontents`` argument,
            the comparison between both files can be either based only on
            file attributes (such as the last modified date), or based on the file contents.

            Args:
                filename1 (str): The 1st file to compare
                filename2 (str): The 2nd file to compare
                comparecontents (bool, optional): When True, the contents of the files are compared. Defaults to False.

            Returns:
                bool: True when the files seem identical

            See Also:
                `SF_FileSystem Help CompareFiles <https://tinyurl.com/ybxpt7eo#CompareFiles>`_
            """
        def CopyFile(
            self, source: str, destination: str, overwrite: bool = ...
        ) -> bool:
            """
            Copies one or more files from one location to another.
            Returns ``True`` if at least one file has been copied or ``False`` if an error occurred.

            Args:
                source (str): FileName or NamePattern which can include wildcard characters, for one or more files to be copied
                destination (str): FileName where the single Source file is to be copied
                    or FolderName where the multiple files from Source are to be copied.
                    If FolderName does not exist, it is created
                    anyway, wildcard characters are not allowed in Destination
                overwrite (bool, optional): If True (default), files may be overwritten. Defaults to True.
                    CopyFile will fail if Destination has the read-only attribute set, regardless of the value of Overwrite.

            Returns:
                bool: True if at least one file has been copied.
                    False if an error occurred.
                    An error also occurs if a source using wildcard characters doesn't match any files.
                    The method stops on the first error it encounters.
                    No attempt is made to roll back or undo any changes made before an error occurs.

            Note:
                - If ``destination`` does not exist, it is created.
                - Wildcard characters are not allowed in ``destination``.

            See Also:
                `SF_FileSystem Help CopyFile <https://tinyurl.com/ybxpt7eo#CopyFile>`_
            """
        def CopyFolder(
            self, source: str, destination: str, overwrite: bool = ...
        ) -> bool:
            """
            Copies one or more folders from one location to another.
            Returns ``True`` if at least one folder has been copied or ``False`` if an error occurred.

            An error will also occur if the ``source`` parameter uses wildcard characters and does
            not match any folders.

            The method stops immediately after it encounters an error.
            The method does not roll back nor does it undo changes made before the error occurred.

            Args:
                source (str): FolderName or NamePattern which can include wildcard characters, for one or more folders to be copied
                destination (str): FolderName where the single Source folder is to be copied
                    or FolderName where the multiple folders from Source are to be copied.
                    If FolderName does not exist, it is created
                    anyway, wildcard characters are not allowed in Destination.
                overwrite (bool, optional): If True (default), folders and their content may be overwritten. Defaults to True.
                    CopyFile will fail if Destination has the read-only attribute set, regardless of the value of Overwrite.

            Note:
                - If ``destination`` does not exist, it is created.
                - Wildcard characters are not allowed in ``destination``.

            Returns:
                bool: True if at least one folder has been copied. False if an error occurred.
                    An error also occurs if a source using wildcard characters doesn't match any folders.
                    The method stops on the first error it encounters.
                    No attempt is made to roll back or undo any changes made before an error occurs.

            See Also:
                `SF_FileSystem Help CopyFolder <https://tinyurl.com/ybxpt7eo#CopyFolder>`_
            """
        def CreateFolder(self, foldername: str) -> bool:
            """
            Creates the specified FolderName. Returns ``True`` if the folder could be successfully created.

            If the specified folder has a parent folder that does not exist, it is created.

            Args:
                foldername (str): a string representing the folder to create. It must not exist

            Returns:
                bool: True if FolderName is a valid folder name, does not exist and creation was successful.
                False otherwise including when FolderName is a file.

            See Also:
                `SF_FileSystem Help CreateFolder <https://tinyurl.com/ybxpt7eo#CreateFolder>`_
            """
        def CreateTextFile(
            self, filename: str, overwrite: bool = ..., encoding: str = ...
        ) -> SFScriptForge.SF_TextStream:
            """
            Creates a specified file and returns a TextStream service instance that can be used to write to the file.

            The method returns ``None`` if an error occurred.

            Args:
                filename (str): The name of the file to be created.
                overwrite (bool, optional): Boolean value that determines if filename can be overwritten (default = True).
                encoding (str, optional): The character set to be used. The default encoding is "UTF-8".

            See Also:
                `SF_FileSystem Help CreateTextFile <https://tinyurl.com/ybxpt7eo#CreateTextFile>`_
            """
        def DeleteFile(self, filename: str) -> bool:
            """
            Deletes one or more files.

            Returns ``True`` if at least one file has been deleted or ``False`` if an error occurred.

            An error will also occur if the ``filename`` parameter uses wildcard characters and does not match any files.

            The files to be deleted must not be readonly.

            The method stops immediately after it encounters an error.
            The method does not roll back nor does it undo changes made before the error occurred.

            Args:
                filename (str): FileName or NamePattern which can include wildcard characters,
                    for one or more files to be deleted.

            Returns:
                bool: True if at least one file has been deleted. False if an error occurred.
                    An error also occurs if a FileName using wildcard characters doesn't match any files.
                    The method stops on the first error it encounters.
                    No attempt is made to roll back or undo any changes made before an error occurs.

            See Also:
                `SF_FileSystem Help DeleteFile <https://tinyurl.com/ybxpt7eo#DeleteFile>`_
            """
        def DeleteFolder(self, foldername: str) -> bool:
            """
            Deletes one or more folders.

            Returns ``True`` if at least one folder has been deleted or ``False`` if an error occurred.

            An error will also occur if the ``foldername`` parameter uses wildcard characters and
            does not match any folders.

            The folders to be deleted must not be readonly.

            The method stops immediately after it encounters an error.
            The method does not roll back nor does it undo changes made before the error occurred.

            Args:
                foldername (str): FolderName or NamePattern which can include wildcard characters, for one or more Folders to be deleted.

            Returns:
                bool: True if at least one folder has been deleted. False if an error occurred.

            See Also:
                `SF_FileSystem Help DeleteFolder <https://tinyurl.com/ybxpt7eo#DeleteFolder>`_
            """
        # 7.4
        # def ExtensionFolder(self, extension: str) -> str:
        #     """
        #     Returns the extension part of a File- or FolderName, without the dot (.).
        #     The method does not check for the existence of the specified file or folder.

        #     Args:
        #         extension (str): Path and file name

        #     Returns:
        #         str: The extension without a leading dot. May be empty
        #     """
        def FileExists(self, filename: str) -> bool:
            """
            Return ``True`` if the given file exists

            Args:
                filename (str): a string representing a file

            Returns:
                bool: True if FileName is a valid File name and it exists.
                    False otherwise including when FileName is a folder

            See Also:
                `SF_FileSystem Help FileExists <https://tinyurl.com/ybxpt7eo#FileExists>`_
            """
        def Files(self, foldername: str, filter: str = ...) -> Tuple[str, ...]:
            """
            Gets a tuple of the FileNames stored in the given folder. The folder must exist

            If the argument ``foldername`` specifies a folder that does not exist, an exception is raised.

            The resulting tuple may be filtered with wildcards.

            Args:
                foldername (str): the folder to explore
                filter (str, optional): contains wildcards ("?" and "*") to limit the list to the relevant files (default = ""). Defaults to "".

            Returns:
                Tuple[str, ...]: An tuple of strings, each entry is the FileName of an existing file
            
            Example:
                .. code::
                
                    >>> a = session.Files("c:\\Windows", "win*.exe")
                    >>> print(a)
                    ('file:///c:/Windows/winhlp32.exe',)

            See Also:
                `SF_FileSystem Help Files <https://tinyurl.com/ybxpt7eo#Files>`_
            """
        def FolderExists(self, foldername: str) -> bool:
            """
            Return ``True`` if the given folder name exists

            If the ``foldername`` parameter is actually an existing file name, the method returns ``False``.

            Args:
                foldername (str): a string representing a folder

            Returns:
                bool: True if FolderName is a valid folder name and it exists.
                    False otherwise including when FolderName is a file.

            See Also:
                `FolderExists <https://tinyurl.com/ybxpt7eo#FolderExists>`_
            """
        def GetBaseName(self, filename: str) -> str:
            """
            Returns the BaseName part of the last component of a File or FolderName, without its extension.
            The method does not check for the existence of the specified file or folder.

            Args:
                filename (str): Path and file name

            Returns:
                str: The BaseName of the given argument in native operating system format. May be empty.

            See Also:
                `SF_FileSystem Help GetBaseName <https://tinyurl.com/ybxpt7eo#GetBaseName>`_
            """
        def GetExtension(self, filename: str) -> str:
            """
            Returns the extension part of a File- or FolderName, without the dot (.).
            The method does not check for the existence of the specified file or folder

            Args:
                filename (str): Path and file name

            Returns:
                str: The extension without a leading dot. May be empty.
            
            Example:
                .. code::

                    >>> print(session.GetExtension("C:\\Windows\\Notepad.exe"))
                    'exe'

            See Also:
                `SF_FileSystem Help GetExtension <https://tinyurl.com/ybxpt7eo#GetExtension>`_
            """
        def GetFileLen(self, filename: str) -> int:
            """
            The builtin ``FileLen`` Basic function returns the number of bytes contained
            in a file as a Long value, i.e. up to 2GB.

            Args:
                filename (str): a string representing a file

            Returns:
                float: File size if FileName exists
            
            Example:
                .. code::

                    >>> print(session.GetFileLen("C:\\Windows\\Notepad.exe"))
                    '201728'

            See Also:
                `SF_FileSystem Help GetFileLen <https://tinyurl.com/ybxpt7eo#GetFileLen>`_
            """
        def GetFileModified(self, filename: str) -> datetime.datetime:
            """
            Returns the last modified date for the given file.

            Args:
                filename (str): a string representing an existing file

            Returns:
                datetime.datetime: The modification date and time.
           
            Example:
                .. code::

                    >>> print(session.GetFileModified("C:\\Windows\\Notepad.exe"))
                    2022-04-03 21:12:26

            See Also:
                `SF_FileSystem Help GetFileModified <https://tinyurl.com/ybxpt7eo#GetFileModified>`_
            """
        def GetName(self, filename: str) -> str:
            """
            Returns the last component of a File- or FolderName.

            The method does not check for the existence of the specified file or folder

            Args:
                filename (str): Path and file name

            Returns:
                str: The last component of the full file name in native operating system format.

            Example:
                .. code-block:: python

                    >>> print(session.GetName("C:\\Windows\\Notepad.exe"))
                   Notepad.exe

            See Also:
                `SF_FileSystem Help GetName <https://tinyurl.com/ybxpt7eo#GetName>`_
            """
        def GetParentFolderName(self, filename: str) -> str:
            """
            Returns a string containing the name of the parent folder of the last component in a specified File- or FolderName.

            The method does not check for the existence of the specified file or folder

            Args:
                filename (str): Path and file name

            Returns:
                str: A FolderName including its final path separator

            Example:
                .. code::

                    >>> print(session.GetParentFolderName("C:\\Windows\\Notepad.exe"))
                   file:///C:/Windows/

            See Also:
                `SF_FileSystem Help GetParentFolderName <https://tinyurl.com/ybxpt7eo#GetParentFolderName>`_
            """
        def GetTempName(self) -> str:
            """
            Returns a randomly generated temporary file name that is useful for performing
            operations that require a temporary file : the method does not create any file.

            Returns:
                str: A FileName as a string. The FileName does not have any suffix.

            Example:
                .. code-block:: python

                    >>> a = f"{session.GetTempName()}.txt"
                    >>> print(a)
                    file:///C:/Users/user/AppData/Local/Temp/SF_419740.txt
            """
        def HashFile(self, filename: str, algorithm: str) -> str:
            """
            Gets a hexadecimal string representing a checksum of the given file.

            Args:
                filename (str): a string representing a file
                algorithm (str): The hashing algorithm to use.

            Returns:
                str: The requested checksum as a string. Hexadecimal digits are lower-cased.
                A zero-length string when an error occurred.

            Note:
                Arg ``algorithm`` support algorithms are:
                    - MD5
                    - SHA1
                    - SHA224
                    - SHA256
                    - SHA384
                    - SHA512

            Example:
                .. code::

                    >>> print(fso.HashFile("C:\\Windows\\Notepad.exe", "MD5"))
                    bbe80313cf12098d3fc4d8a42e9dbb33

            See Also:
                `SF_FileSystem Help HashFile <https://tinyurl.com/ybxpt7eo#HashFile>`_
            """
        def MoveFile(self, source: str, destination: str) -> bool:
            """
            Moves one or more files from one location to another.

            Returns ``True`` if at least one file has been moved or ``False`` if an error occurred.

            An error will also occur if the source parameter uses wildcard characters and
            does not match any files.

            The method stops immediately after it encounters an error. The method does not
            roll back nor does it undo changes made before the error occurred.

            Args:
                source (str): FileName or NamePattern which can include wildcard characters, for one or more files to be moved
                destination (str): FileName where the single Source file is to be moved.
                    If Source and Destination have the same parent folder MoveFile amounts to renaming the Source or
                    FolderName where the multiple files from Source are to be moved.
                    If FolderName does not exist, it is created anyway, wildcard characters are not allowed in Destination.
                    

            Returns:
                bool: True if at least one file has been moved. False if an error occurred.

            See Also:
                `SF_FileSystem Help MoveFile <https://tinyurl.com/ybxpt7eo#MoveFile>`_
            """
        def MoveFolder(self, source: str, destination: str) -> bool:
            """
            Moves one or more folders from one location to another.

            Returns ``True`` if at least one folder has been moved or ``False`` if an error occurred.

            An error will also occur if the source parameter uses wildcard characters and does
            not match any folders.

            The method stops immediately after it encounters an error. The method does not roll
            back nor does it undo changes made before the error occurred.

            Args:
                source (str): FolderName or NamePattern which can include wildcard characters, for one or more folders to be moved
                destination (str): FolderName where the single Source folder is to be moved. FolderName must not exist or
                    FolderName where the multiple folders from Source are to be moved.
                    If FolderName does not exist, it is created anyway, wildcard characters are not allowed in Destination.

            Returns:
                bool: True if at least one folder has been moved. False if an error occurred.

            See Also:
                `SF_FileSystem Help MoveFolder <https://tinyurl.com/ybxpt7eo#MoveFolder>`_
            """
        def OpenTextFile(
            self,
            filename: str,
            iomode: int = ...,
            create: bool = ...,
            encoding: str = ...,
        ) -> SFScriptForge.SF_TextStream:
            """
            Opens a specified file and returns a TextStream object that can be used to read from, write to, or append to the file.

            Args:
                filename (str): Identifies the file to open
                iomode (int, optional): Indicates input/output mode.
                create (bool, optional): Boolean value that indicates whether a new file can be created if the specified filename doesn't exist.
                    The value is True if a new file and its parent folders may be created; False if they aren't created. Defaults to False.
                encoding (str, optional):  The character set that should be used. Defaults to "UTF-8".

            Returns:
                SFScriptForge.SF_TextStream: SF_TextStream instance representing the opened file
                    or None if an error occurred. The method does not check if the file is really
                    a text file. It doesn't check either if the given encoding is implemented in
                    LibreOffice nor if it is the right one.

            Note:
                Arg ``iomode`` Constants.
                    - SF_FileSystem.ForReading (Default)
                    - SF_FileSystem.ForWriting
                    - SF_FileSystem.ForAppending

                Arg ``encoding`` is one of the following `Character Sets <https://www.iana.org/assignments/character-sets/character-sets.xhtml>`_
                
                LibreOffice does not implement all existing encoding sets.

            Example:
                .. code::

                    >>> a = session.OpenTextFile("c:\\temp\\test.txt", session.ForReading, False)
                    >>> if a is not None:
                            a.ReadAll()
                    'ScriptForge Rocks\\r\\n'
                    >>> a.CloseFile()
                    True

            See Also:
                `SF_FileSystem Help OpenTextFile <https://tinyurl.com/ybxpt7eo#OpenTextFile>`_
            """
        def PickFile(
            self, defaultfile: str = ..., mode: str = ..., filter: str = ...,
        ) -> str:
            """
            Opens a dialog box to open or save files.

            If the ``SAVE`` mode is set and the picked file exists, a warning message will be displayed.
            

            Args:
                defaultfile (str, optional): This argument is a string composed of a folder and file name.
                mode (str, optional): A string value that can be either "OPEN" (for input files) or
                    "SAVE" (for output files). The default value is "OPEN".
                filter (str, optional): The extension of the files displayed when the dialog is opened (default = no filter).

            Note:
                ``defaultfile`` notes:
                    - The folder part indicates the folder that will be shown when the dialog opens (default = the last selected folder).
                    - The file part designates the default file to open or save.

            Returns:
                str: The selected FileName in URL format or "" if the dialog was cancelled.

            See Also:
                `SF_FileSystem Help PickFile <https://tinyurl.com/ybxpt7eo#PickFile>`_
            """
        def PickFolder(self, defaultfolder: str = ..., freetext: str = ...) -> str:
            """
            Display a FolderPicker dialog box.

            Args:
                defaultfolder (str, optional): the FolderName from which to start. Default = the last selected folder. Defaults to ScriptForge.cstSymEmpty.
                freetext (str, optional): text to display in the dialog. Defaults to "".

            Returns:
                str: The selected FolderName in URL or operating system format.
                A zero-length string if the dialog was cancelled.

            See Also:
                `SF_FileSystem Help PickFolder <https://tinyurl.com/ybxpt7eo#PickFolder>`_
            """
        def SubFolders(self, foldername: str, filter: str = ...) -> Tuple[str, ...]:
            """
            Returns a zero-based array of strings corresponding to the folders stored
            in a given ``foldername``.

            The list may be filtered with wildcards.

            Args:
                foldername (str): A string representing a folder. The folder must exist.
                    foldername must not designate a file.
                filter (str, optional): A string containing wildcards ("?" and "*") that will
                    be applied to the resulting list of folders (default = "").

            Returns:
                tuple: tuple of strings corresponding to the folders stored in a given foldername.

            See Also:
                `SF_FileSystem Help SubFolders <https://tinyurl.com/ybxpt7eo#SubFolders>`_
            """
        @classmethod
        def _ConvertFromUrl(cls, filename: str) -> str: ...
        # _ConvertFromUrl Alias for same function in FileSystem Basic module
        # endregion Methods
        
        # region Properties
        @property
        def FileNaming(self) -> str:
            """
            Gets/Sets the current files and folders notation, either ANY, URL or SYS:
            
            - "ANY": (default) the methods of the FileSystem service accept both URL
                and current operating system's notation for input arguments but always return URL strings.
            - "URL": the methods of the FileSystem service expect URL notation for input arguments
                and return URL strings.
            - "SYS": the methods of the FileSystem service expect current operating system's notation
                for both input arguments and return strings.

            Once set, the FileNaming property remains unchanged either until the end of the
            LibreOffice session or until it is set again.
            """
        @property
        def ConfigFolder(self) -> str:
            """Gets the configuration folder of LibreOffice."""
        @property
        def ExtensionsFolder(self) -> str:
            """Gets the folder where extensions are installed."""
        @property
        def HomeFolder(self) -> str:
            """Gets the user home folder."""
        @property
        def InstallFolder(self) -> str:
            """Gets the installation folder of LibreOffice."""
        @property
        def TemplatesFolder(self) -> str:
            """Gets the folder containing the system templates files."""
        @property
        def TemporaryFolder(self) -> str:
            """
            Gets the temporary files folder defined in the LibreOffice path settings.
            """
        @property
        def UserTemplatesFolder(self) -> str:
            """Gets the folder containing the user-defined template files."""
        # endregion Properties
    # endregion SF_FileSystem CLASS
    
    # region SF_L10N CLASS
    class SF_L10N(SFServices):
        """
        This service provides a number of methods related to the translation of strings
        with minimal impact on the program's source code.
        The methods provided by the L10N service can be used mainly to:
            Create POT files that can be used as templates for translation of all strings in the program.
            Get translated strings at runtime for the language defined in the Locale property.

        See Also:
            `ScriptForge.L10N service <https://tinyurl.com/y77mbtp9>`_
        """

        @classmethod
        def ReviewServiceArgs(
            cls, foldername: str = ..., locale: str = ..., encoding: str = ...
        ) -> Tuple[str, str, str]:
            """
            Transform positional and keyword arguments into positional only
            """
        def AddText(
            self, context: str = ..., msgid: str = ..., comment: str = ...
        ) -> bool:
            """
            Adds a new entry in the list of localizable strings. It must not exist yet.

            The method returns ``True`` if successful.

            Args:
                context (str, optional): The key to retrieve the translated string with the GetText method.
                    This parameter has a default value of "".
                msgid (str, optional): The untranslated string, which is the text appearing in the program code.
                comment (str, optional): Optional comment to be added alongside the string to help translators.

            Note:
                ``msgid`` must not be empty. The ``msgid`` becomes the key to retrieve the translated
                string via GetText method when context is empty.

                The ``msgid`` string may contain any number of placeholders (%1 %2 %3 ...) for dynamically
                modifying the string at runtime.

            Returns:
                bool: True if successful.

            See Also:
                `SF_L10N Help AddText <https://tinyurl.com/y77mbtp9#AddText>`_
            """
        def AddTextsFromDialog(self, dialog: object) -> bool:
            """
            Add all fixed text strings of a dialog to the list of localizable text strings.
            
            Added texts are:
                - the title of the dialog
                - the caption associated with next control types: Button, CheckBox, FixedLine, FixedText, GroupBox and RadioButton
                - the content of list- and comboboxes
                - the tip- or helptext displayed when the mouse is hovering the control
            
            The current method has method SFDialogs.SF_Dialog.GetTextsFromL10N as counterpart.
            The targeted dialog must not be open when the current method is run.

            Args:
                dialog (object): a SFDialogs.Dialog service instance

            Returns:
                bool: True when successful.

            See Also:
                `SF_L10N Help AddTextsFromDialog <https://tinyurl.com/y77mbtp9#AddTextsFromDialog>`_
            """
        def ExportToPOTFile(
            self, filename: str, header: str = ..., encoding: str = ...
        ) -> bool:
            """
            Export a set of untranslated strings as a POT file.
            The set of strings has been built either by a succession of AddText() methods
            or by a successful invocation of the L10N service with the FolderName argument.
            The generated file should pass successfully the "msgfmt --check 'the pofile'" GNU command.

            Args:
                filename (str): the complete file name to export to. If it exists, is overwritten without warning.
                header (str, optional): Comments that will appear on top of the generated file. Do not include any leading "#".
                    If the string spans multiple lines, insert escape sequences (\\n) where relevant.
                    A standard header will be added anyway. Defaults to "".
                encoding (str, optional):  The character set that should be used. Defaults to "UTF-8".

            Returns:
                bool: True if successful.
            
            Note:
                ``encoding`` is one of the following `Character Sets <https://www.iana.org/assignments/character-sets/character-sets.xhtml>`_
                
                LibreOffice does not implement all existing encoding sets.

            See Also:
                `SF_L10N Help ExportToPOTFile <https://tinyurl.com/y77mbtp9#ExportToPOTFile>`_
            """
        def GetText(self, msgid: Any, *args: Any) -> str:
            """
            Gets the translated string corresponding to the given ``msgid`` argument.

            A list of arguments may be specified to replace the placeholders (%1, %2, ...)
            in the string.

            If no translated string is found, the method returns the untranslated string
            after replacing the placeholders with the specified arguments.

            Args:
                msgid (Any): The untranslated string, which is the text appearing
                    in the program code. It must not be empty. It may contain any
                    number of placeholders (%1 %2 %3 ...) that can be used to
                    dynamically insert text at runtime.
                args (Any): Values to be inserted into the placeholders. Any variable type is allowed, however only strings, numbers and dates will be considered.

            Returns:
                str: The translated string. If not found the MsgId string or the Context string
                anyway the substitution is done

            Note:
                Besides using a single ``msgid`` string, this method also accepts the following formats:
                    - the untranslated text (MsgId)
                    - the reference to the untranslated text (Context)
                    - both (Context|MsgId) : the pipe character is essential

            See Also:
                `SF_L10N Help GetText <https://tinyurl.com/y77mbtp9#GetText>`_
            """
        # endregion Methods
        
        # region Properties
        @property
        def Folder(self) -> str:
            """
            Gets The folder containing the PO files (see the FileSystem.FileNaming property to learn about the notation used).
            """
        @property
        def Languages(self) -> Tuple[str, ...]:
            """
            A listing all the base names (without the ".po" extension)
            of the PO-files found in the specified Folder.
            """
        @property
        def Locale(self) -> str:
            """
            Gets The currently active language-COUNTRY combination.
            This property will be initially empty if the service was instantiated without
            any of the optional arguments.
            """
        # endregion Properties
    # endregion SF_L10N CLASS
    
    # region SF_Platform CLASS
    class SF_Platform(SFServices, metaclass=_Singleton):
        """
        The 'Platform' service implements a collection of properties about the actual execution environment
        and context :
            the hardware platform
            the operating system
            the LibreOffice version
            the current user
        All those properties are read-only.
        The implementation is mainly based on the 'platform' module of the Python standard library.

        See Also:
            `ScriptForge.Platform service <https://tinyurl.com/ybuvx3v4>`_
        """
        # Python helper functions
        py: str = ...
        # region Properties
        @property
        def Architecture(self) -> str:
            """Gets the actual bit architecture"""
        @property
        def ComputerName(self) -> str:
            """Gets the computer's network name"""
        @property
        def CPUCount(self) -> int:
            """Gets the number of Central Processor Units"""
        @property
        def CurrentUser(self) -> str:
            """Gets the name of logged in user"""
        @property
        def Fonts(self) -> Tuple[str, ...]:
            """
            Gets tuple of strings containing the names of all available fonts.
            """
        @property
        def Locale(self) -> str:
            """
            Returns the operating system locale as a string in the format language-COUNTRY (la-CO).
            
            Examples: "en-US", "pt-BR", "fr-BE".
            """
        @property
        def Machine(self) -> str:
            """Gets the machine type like 'i386' or 'x86_64'"""
        @property
        def OfficeVersion(self) -> str:
            """
            The actual LibreOffice version expressed as 'LibreOffice w.x.y.z (The Document Foundation)'.
            
            Example: 'LibreOffice 7.1.1.2 (The Document Foundation, Debian and Ubuntu)
            """
        @property
        def OSName(self) -> str:
            """Gets the name of the operating system like 'Linux' or 'Windows'"""
        @property
        def OSPlatform(self) -> str:
            """
            Gets a single string identifying the underlying platform with as much useful
            and human-readable information as possible.
            
            Such as ``Linux-4.15.0-117-generic-x86_64-with-Ubuntu-18.04-bionic``
            """
        @property
        def OSRelease(self) -> str:
            """Gets the operating system's release such as ``4.15.0-117-generic``"""
        @property
        def OSVersion(self) -> str:
            """
            Gets the name of the operating system build or version.
            
            Such as ``118-Ubuntu SMP Fri Sep 4 20:02:41 UTC 2020``
            """
        @property
        def Printers(self) -> Tuple[str]:
            """
            The tuple of available printers.

            The default printer is put in the first position (index = 0).
            """
        @property
        def Processor(self) -> str:
            """
            Gets the (real) processor name, e.g. 'amdk6'. Might return the same value as Machine.
            """
        @property
        def PythonVersion(self) -> str:
            """
            Gets the Python version as string 'Python major.minor.patchlevel'.
            
            Such as ``Python 3.7.7``
            """
        # endregion Properties
    # endregion SF_Platform CLASS
    
    class SF_Session(SFServices, metaclass=_Singleton):
        """
        The Session service gathers various general-purpose methods about:
        - UNO introspection
        - the invocation of external scripts or programs

        See Also:
            `ScriptForge.Session service <https://tinyurl.com/yaf7co37>`_
        """
        # region CONST

        # Class constants               Where to find an invoked library ?
        SCRIPTISEMBEDDED: Literal["document"]  # in the document
        SCRIPTISAPPLICATION: Literal["application"]  # in any shared library (Basic)
        SCRIPTISPERSONAL: Literal["user"]  # in My Macros (Python)
        SCRIPTISPERSOXT: Literal[
            "user:uno_packages"
        ]  # in an extension installed for the current user (Python)
        SCRIPTISSHARED: Literal["share"]  # in LibreOffice macros (Python)
        SCRIPTISSHAROXT: Literal[
            "share:uno_packages"
        ]  # in an extension installed for all users (Python)
        SCRIPTISOXT: Literal[
            "uno_packages"
        ]  # in an extension but the installation parameters are unknown (Python)
        # endregion CONST
        
        # region Methods
        @classmethod
        def ExecuteBasicScript(
            cls, scope: str = ..., script: str = ..., *args: Any
        ) -> Any:
            """
            Execute the Basic script given its name and location and fetch its result if any.

            If the script returns nothing, which is the case of procedures defined with Sub,
            the returned value is Empty.

            Args:
                scope (str, optional): "String specifying where the script is stored.
                script (str, optional): String specifying the script to be called in the format
                    "library.module.method" as a case-sensitive string.
                args (any, optional): The arguments to be passed to the called script.

            Note:
                Arg ``scope`` can be either:
                    - "document" (constant ``session.SCRIPTISEMBEDDED``)
                    - "application" (constant ``session.SCRIPTISAPPLICATION``).

                Arg ``script`` can be:
                    - The library is loaded in memory if necessary.
                    - The module must not be a class module.
                    - The method may be a ``Sub`` or a ``Function``.

            Returns:
                Any: The value returned by the call to the script
            
            See Also:
                `ExecuteBasicScript <https://tinyurl.com/yaf7co37#ExecuteBasicScript>`_

                `Scripting Framework URI Specification <https://wiki.openoffice.org/wiki/Documentation/DevGuide/Scripting/Scripting_Framework_URI_Specification>`_
            """
        @classmethod
        def ExecuteCalcFunction(cls, calcfunction: str, *args: Any) -> Any:
            """
            Execute a Calc function using its English name and based on the given arguments.
            If the arguments are arrays, the function is executed as an
            `array formula <https://tinyurl.com/5xtjus4a>`_.

            Args:
                calcfunction (str): The name of the Calc function to be called, in English.
                args (any, optional): The arguments to be passed to the called Calc function.
                    Each argument must be either a string, a numeric value or an array of arrays
                    combining those types.

            Example:
                .. code::

                    >>> session.ExecuteCalcFunction("AVERAGE", 1, 5, 3, 7) # 4
                    >>> session.ExecuteCalcFunction("ABS", ((-1, 2, 3), (4, -5, 6), (7, 8, -9)))[2][2] # 9
                    >>> session.ExecuteCalcFunction("LN", -3)

            Returns:
                Any: results of function.

            See Also:
                `ExecuteCalcFunction <https://tinyurl.com/yaf7co37#ExecuteCalcFunction>`_
            """
        @classmethod
        def ExecutePythonScript(
            cls, scope: str = ..., script: str = ..., *args: Any
        ) -> Any:
            """
            Execute the Python script given its location and name, fetch its result if any.
            Result can be a single value or an array of values.

            If the script is not found, or if it returns nothing, the returned value is Empty.

            The LibreOffice Application Programming Interface (API) Scripting Framework
            supports inter-language script execution between Python and Basic, or other
            supported programming languages for that matter. Arguments can be passed back and
            forth across calls, provided that they represent primitive data types that both
            languages recognize, and assuming that the Scripting Framework converts them appropriately.


            Args:
                scope (str, optional): One of the applicable constants listed. See Note
                script (str, optional): Either "library/module.py$method" or "module.py$method"
                    or "myExtension.oxt|myScript|module.py$method" as a case-sensitive string.

            Note:
                Arg ``scope`` `constants <https://tinyurl.com/yaf7co37#constants>`_:
                    - SCRIPTISEMBEDDED
                    - SCRIPTISAPPLICATION
                    - SCRIPTISPERSONAL
                    - SCRIPTISPERSOXT
                    - SCRIPTISSHARED
                    - SCRIPTISSHAROXT
                    - SCRIPTISOXT

                Arg ``script`` values:
                    - library: The folder path to the Python module.
                    - myScript: The folder containing the Python module.
                    - module.py: The Python module.
                    - method: The Python function.

            Returns:
                Any: The value(s) returned by the call to the script. If > 1 values, enclosed in a tuple.
            
            See Also:
                `ExecutePythonScript <https://tinyurl.com/yaf7co37#ExecutePythonScript>`_

                `Scripting Framework URI Specification <https://wiki.openoffice.org/wiki/Documentation/DevGuide/Scripting/Scripting_Framework_URI_Specification>`_
            """
        def HasUnoMethod(self, unoobject: object, methodname: str) -> bool:
            """
            Returns True if a UNO object contains the given method.
            Code-snippet derived from XRAY

            Args:
                unoobject (object): the object to identify
                methodname (str): the name of the method as a string. The search is case-sensitive

            Returns:
                bool: False when the method is not found or when an argument is invalid
            """
        def HasUnoProperty(self, unoobject: object, propertyname: str) -> bool:
            """
            Returns True if a UNO object contains the given property.
            Code-snippet derived from XRAY

            Args:
                unoobject (object): the object to identify
                propertyname (str): the name of the property as a string. The search is case-sensitive

            Returns:
                bool: False when the property is not found or when an argument is invalid

            See Also:
                `HasUnoProperty <https://tinyurl.com/yaf7co37#HasUnoProperty>`_
            """
        @classmethod
        def OpenURLInBrowser(cls, url: str) -> None:
            """
            Opens a URL in the default browser.

            Args:
                url (str): The URL to open in the browser

            See Also:
                `OpenURLInBrowser <https://tinyurl.com/yaf7co37#OpenURLInBrowser>`_
            """
        def RunApplication(self, command: Any, parameters: str) -> bool:
            """
            Executes an arbitrary system command

            Args:
                command (Any):  The command to execute
                    This may be an executable file or a document which is registered with an application
                    so that the system knows what application to launch for that document
                parameters (str): a list of space separated parameters as a single string.
                    The method does not validate the given parameters, but only passes them to the specified command

            Returns:
                bool: True if success

            Example:
                .. code::
                
                    session.RunApplication"Notepad.exe")
                    session.RunApplication("C:\myFolder\myDocument.odt")
                    session.RunApplication("kate", "/home/me/install.txt") # Linux

            See Also:
                `RunApplication <https://tinyurl.com/yaf7co37#RunApplication>`_
            """
        def SendMail(
            self,
            recipient: str,
            cc: str = ...,
            bcc: str = ...,
            subject: str = ...,
            body: str = ...,
            filenames: str = ...,
            editmessage:bool=...,
        ) -> None:
            """
            Send a message (with or without attachments) to recipients from the user's mail client.
            The message may be edited by the user before sending or, alternatively, be sent immediately

            Args:
                recipient (str): an email addresses (To recipient)  
                cc (str, optional): a comma-delimited list of email addresses (carbon copy). Defaults to ''.
                bcc (str, optional): a comma-delimited list of email addresses (blind carbon copy). Defaults to ''.
                subject (str, optional): the header of the message. Defaults to ''.
                body (str, optional): the unformatted text of the message. Defaults to ''.
                filenames (str, optional): a comma-separated list of filenames to attach to the mail. SF_FileSystem naming conventions apply. Defaults to ''.
                editmessage (bool, optional): when True (default) the message is editable before being sent. Defaults to True.

            See Also:
                `SendMail <https://tinyurl.com/yaf7co37#SendMail>`_
            """
        def UnoObjectType(self, unoobject: object) -> str:
            """
            Identify the UNO type of an UNO object.
            Code-snippet derived from XRAY.

            Args:
                unoobject (object):  the object to identify

            Returns:
                str: com.sun.star. ... as a string or
                a zero-length string if identification was not successful

            See Also:
                `UnoObjectType <https://tinyurl.com/yaf7co37#UnoObjectType>`_
            """
        def UnoMethods(self, unoobject: object) -> Tuple[str, ...]:
            """
            Returns a tuple of the methods callable from an UNO object.

            Args:
                unoobject (object):  the object to identify

            Returns:
                Tuple[str, ...]: A zero-based sorted tuple. May be empty

            See Also:
                `UnoMethods <https://tinyurl.com/yaf7co37#UnoMethods>`_
            """
        def UnoProperties(self, unoobject: object) -> Tuple[str, ...]:
            """
            Returns a tuple of the properties of an UNO object.

            Args:
                unoobject (object):  the object to identify

            Returns:
                Tuple[str, ...]: A zero-based sorted tuple. May be empty

            See Also:
                `UnoProperties <https://tinyurl.com/yaf7co37#UnoProperties>`_
            """
        def WebService(self, uri: str) -> str:
            """
            Get some web content from a URI

            Args:
                uri (str): text of the web service

            Returns:
                str: The web page content of the URI

            Example:
                .. code::

                    session.WebService(
                        "wiki.documentfoundation.org/api.php?hidebots=1&days=7&limit=50&action=feedrecentchanges&feedformat=rss"
                        )

            See Also:
                `WebService <https://tinyurl.com/yaf7co37#WebService>`_
            """
        # endregion Methods
    # endregion SF_Session CLASS
    
    # region SF_String CLASS
    class SF_String(SFServices, metaclass=_Singleton):
        """
        Focus on string manipulation, regular expressions, encodings and hashing algorithms.
        The methods implemented in Basic that are redundant with Python builtin functions
        are not duplicated

        See Also:
            `ScriptForge.String service <https://tinyurl.com/y9hm6agu>`_
        """
        # region Methods
        @classmethod
        def HashStr(cls, inputstr: str, algorithm: str) -> str:
            """
            Return a hexadecimal string representing a checksum of the given input string

            Args:
                inputstr (str): the string to be hashed
                algorithm (str): The hashing algorithm to use.

            Returns:
                str: The requested checksum as a string. Hexadecimal digits are lower-cased.
                A zero-length string when an error occurred
            Note:
                Arg ``algorithm`` support algorithms are:
                    - MD5
                    - SHA1
                    - SHA224
                    - SHA256
                    - SHA384
                    - SHA512

            Example:
                .. code-block:: python

                    SF_String.HashStr("œ∑¡™£¢∞§¶•ªº–≠œ∑´®†¥¨ˆøπ“‘åß∂ƒ©˙∆˚¬", "MD5")	# 616eb9c513ad07cd02924b4d285b9987

            See Also:
                `SF_String Help HashStr <https://tinyurl.com/y9hm6agu#HashStr>`_
            """
        def IsADate(self, inputstr: str, dateformat: str = ...) -> bool:
            """
            Return True if the string is a valid date respecting the given format

            Args:
                inputstr (str): the input string
                dateformat (str, optional): either YYYY-MM-DD (default), DD-MM-YYYY or MM-DD-YYYY. Defaults to 'YYYY-MM-DD'.
                    The dash (-) may be replaced by a dot (.), a slash (/) or a space.

            Returns:
                bool: True if the string contains a valid date and there is at least one character.
                False otherwise or if the date format is invalid
            
            Example:
                .. code::
                
                    >>> print(SF_String.IsADate("2019-12-31", "YYYY-MM-DD"))
                    True
            See Also:
                `SF_String Help IsADate <https://tinyurl.com/y9hm6agu#IsADate>`_
            """
        def IsEmail(self, inputstr: str) -> bool:
            """
            Return True if the string is a valid email address

            Args:
                inputstr (str): the input string    

            Returns:
                bool: True if the string contains an email address and there is at least one character, False otherwise.
            
            Example:
                .. code::
                
                    >>> print(SF_String.IsEmail("first.last@something.org"))
                    True

            See Also:
                `SF_String Help IsEmail <https://tinyurl.com/y9hm6agu#IsEmail>`_
            """
        def IsFileName(self, inputstr: str, osname: str = ...) -> bool:
            """
            Return True if the string is a valid filename in a given operating system

            Args:
                inputstr (str): the input string
                osname (str, optional): ``Windows``, ``Linux``, ``macOS`` or ``Solaris``. Defaults to ScriptForge.cstSymEmpty.
                    The default is the current operating system on which the script is run.

            Returns:
                bool: True if the string contains a valid filename and there is at least one character, False otherwise
            
            Example:
                .. code-block:: python
                
                    >>> print(SF_String.IsFileName("/home/a file name.odt", "Linux"))
                    True
            """
        def IsIBAN(self, inputstr: str) -> bool:
            """
            Returns True if the input string is a valid International Bank Account Number.
            
            Read `International Bank Account Number <https://en.wikipedia.org/wiki/International_Bank_Account_Number>`_

            Args:
                inputstr (str): the input string

            Returns:
                bool: True if the string contains a valid IBAN number. The comparison is not case-sensitive
            
            Example:
                .. code::
                
                    >>> print(SF_String.IsIBAN("BR15 0000 0000 0000 1093 2840 814 P2"))
                    True

            See Also:
                `SF_String Help IsIBAN <https://tinyurl.com/y9hm6agu#IsIBAN>`_
            """
        def IsIPv4(self, inputstr: str) -> bool:
            """
            Return True if the string is a valid IPv4 address

            Args:
                inputstr (str): the input string

            Returns:
                bool: True if the string contains a valid IPv4 address and there is at least one character, False otherwise
            
            Example:
                .. code::
                
                    >>> print(SF_String.IsIPv4("192.168.1.50"))
                    True

            See Also:
                `SF_String Help IsIPv4 <https://tinyurl.com/y9hm6agu#IsIPv4>`_
            """
        def IsLike(
            self, inputstr: str, pattern: str, casesensitive: bool = ...
        ) -> bool:
            """
            Returns True if the whole input string matches a given pattern containing wildcards

            Args:
                inputstr (str): the input string
                pattern (str):  the pattern as a string.
                casesensitive (bool, optional): case sensitive. Defaults to False.

            Note:
                Arg ``pattern`` wildcards:
                    - the ``?`` represents any single character
                    - the ``*`` represents zero, one, or multiple characters

            Returns:
                bool: True if a match is found.
                Zero-length input or pattern strings always return False
            
            Example:
                .. code::
                
                    >>> print(SF_String.IsLike("aAbB", "?A*"))
                    True
                    >>> print(	SF_String.IsLike("C:\\a\\b\\c\\f.odb", "?:*.*"))
                    True

            See Also:
                `SF_String Help IsLike <https://tinyurl.com/y9hm6agu#IsLike>`_
            """
        def IsSheetName(self, inputstr: str) -> bool:
            """
            Return True if the input string can serve as a valid Calc sheet name.
            
            The sheet name must not contain the characters ``[ ] * ? : / \``
            or the character ' (apostrophe) as first or last character.

            Args:
                inputstr (str): the input string

            Returns:
                bool: True if the string is validated as a potential Calc sheet name, False otherwise
            
            Example:
                .. code::
                
                    >>> print(SF_String.IsSheetName('1àbc + "def"'))
                    True

            See Also:
                `SF_String Help IsSheetName <https://tinyurl.com/y9hm6agu#IsSheetName>`_
            """
        def IsUrl(self, inputstr: str) -> bool:
            """
            Return True if the string is a valid absolute URL (Uniform Resource Locator).
            
            The parsing is done by the ParseStrict method of the URLTransformer UNO service.
            `XURLTransformer Interface <https://tinyurl.com/5yhwekmf>`_

            Args:
                inputstr (str): the input string    

            Returns:
                bool: True if the string contains a URL and there is at least one character, False otherwise
            
            Example:
                .. code::
                
                    >>> print(SF_String.IsUrl("http://foo.bar/?q=Test%20URL-encoded%20stuff"))
                    True

            See Also:
                `SF_String Help IsUrl <https://tinyurl.com/y9hm6agu#IsUrl>`_
            """
        def SplitNotQuoted(
            self,
            inputstr: str,
            delimiter: str = ...,
            occurrences: int = ...,
            quotechar: str = ...,
        ) -> tuple:
            """
            Split a string on Delimiter into an array. If Delimiter is part of a quoted (sub)string, it is ignored.
            (used f.i. for parsing of csv-like records)

            Args:
                inputstr (str): the input string.
                    Might contain quoted substrings:
                        The quoting character must be the double quote (")
                delimiter (str, optional): A string of one or more characters that is used to delimit the input string. Defaults to ' '.
                    The default is the space character
                occurrences (int, optional): The number of substrings to return (Default = 0, meaning no limit). Defaults to 0.
                quotechar (str, optional): The quoting character, either " (default) or '. Defaults to '"'.

            Returns:
                tuple: A tuple whose items are chunks of the input string, Delimiter not included
            
            Example:
                .. code::
                
                    >>> print(SF_String.SplitNotQuoted("abc def ghi"))
                    ["abc", "def", "ghi"]
                    >>> print(SF_String.SplitNotQuoted("abc,'"def,ghi"', ","))
                    ["abc", '"def,ghi"']

            See Also:
                `SF_String Help SplitNotQuoted <https://tinyurl.com/y9hm6agu#SplitNotQuoted>`_
            """
        def Wrap(self, inputstr, width=..., tabsize=...) -> Tuple[str, ...]:
            """
            Wraps every single paragraph in text (a string) so every line is at most Width characters long.

            Args:
                inputstr (_type_): the input string.
                width (int, optional): the maximum number of characters in each line. Defaults to 70.
                tabsize (int, optional): before wrapping the text, the existing TAB (Chr(9)) characters are replaced with spaces.
                    TabSize defines the TAB positions at TabSize + 1, 2 * TabSize + 1 , ... N * TabSize + 1.
                    Defaults to 8.

            Returns:
                Tuple[str, ...]: Returns a zero-based tuple of output lines, without final newlines except the pre-existing line-breaks.
                    Tabs are expanded. Symbolic line breaks are replaced by their hard equivalents.
                    If the wrapped output has no content, the returned tuple is empty.

            See Also:
                `SF_String Help Wrap <https://tinyurl.com/y9hm6agu#Wrap>`_
            """
        # endregion Methods
    
        # region Properties
        @property
        def sfCR(self) -> str:
            """Gets Carriage return: Chr(13)"""
        @property
        def sfCRLF(self) -> str:
            """Gets Carriage return + Linefeed: Chr(13) & Chr(10)"""
        @property
        def sfLF(self) -> str:
            """Gets Linefeed: Chr(10)"""
        @property
        def sfNEWLINE(self) -> str:
            """
            Gets Carriage return + Linefeed, which can be
            
            1) Chr(13) & Chr(10) or
            
            2) Linefeed: Chr(10)
            
            depending on the operating system.
            """
        @property
        def sfCR(self) -> str:
            """Gets Horizontal tabulation: Chr(9)"""
        # endregion Properties
    
    # endregion SF_String CLASS
    
    # region SF_TextStream CLASS
    class SF_TextStream(SFServices):
        """
        The TextStream service is used to sequentially read from and write to files opened or created
        using the ScriptForge.FileSystem service..

        See Also:
            `ScriptForge.TextStream service <https://tinyurl.com/y9ubyoel>`_
        """
        # region Methods
        def CloseFile(self) -> bool:
            """
            Empties the output buffer if relevant. Closes the actual input or output stream.

            Returns:
                bool: True if the closure was successful.

            See Also:
                `SF_TextStream Help CloseFile <https://tinyurl.com/y9ubyoel#CloseFile>`_
            """
        def ReadAll(self) -> str:
            """
            Returns all the remaining lines in the text stream as one string. Line breaks are NOT removed.
            The resulting string can be split in lines either by using the usual Split Basic builtin function if the line delimiter is known
            or with the SF_String.SplitLines method.

            Returns:
                str: The read lines. The string may be empty.
            
            Note:
                Line property in incremented only by 1
            """
        def ReadLine(self) -> str:
            """
            Returns the next line in the text stream as a string. Line breaks are removed.
            
            Returns:
                str: The read line. The string may be empty.

            See Also:
                `SF_TextStream Help ReadLine <https://tinyurl.com/y9ubyoel#ReadLine>`_
            """
        def SkipLine(self) -> None:
            """
            Skips the next line when reading a TextStream file.

            See Also:
                `SF_TextStream Help SkipLine <https://tinyurl.com/y9ubyoel#SkipLine>`_
            """
        def WriteBlankLines(self, lines: int) -> None:
            """
            Writes a number of empty lines in the output stream.

            Args:
                lines (int): the number of lines to write

            See Also:
                `SF_TextStream Help WriteBlankLines <https://tinyurl.com/y9ubyoel#WriteBlankLines>`_
            """
        def WriteLine(self, line: str) -> None:
            """
            Writes the given line to the output stream. A newline is inserted if relevant

            Args:
                line (str): the line to write, may be empty.

            See Also:
                `SF_TextStream Help WriteLine <https://tinyurl.com/y9ubyoel#WriteLine>`_
            """
        # endregion Methods

        # region Properties
        @property
        def AtEndOfStream(self) -> bool:
            """
            In reading mode, True indicates that the end of the file has been reached
            In write and append modes, or if the file is not ready => always True
            The property should be invoked BEFORE each ReadLine() method:
            A ReadLine() executed while AtEndOfStream is True will raise an error
            """
        atEndOfStream, atendofstream = AtEndOfStream, AtEndOfStream
        @property
        def Encoding(self) -> str:
            """
            Gets the character set to be used. The default encoding is ``UTF-8``.
            """
        @property
        def FileName(self) -> str:
            """
            Gets the name of the current file either in URL format or in the native
            operating system's format, depending on the current value of the FileNaming
            property of the FileSystem service.
            """
        @property
        def IOMode(self) -> str:
            """
            Gets the input/output mode. Possible values are READ, WRITE or APPEND.
            """
        @property
        def Line(self) -> int:
            """
            Gets the number of lines read or written so far.
            """
        @property
        def NewLine(self) -> str:
            """
            Gets/Sets the current delimiter to be inserted between two successive written lines.
            The default value is the native line delimiter in the current operating system.
            """
        # endregion Properties

    # endregion SF_TextStream CLASS

    # region SF_Timer CLASS
    class SF_Timer(SFServices):
        """
        The Timer service measures the amount of time it takes to run user scripts.

        A Timer measures durations. It can be:
            - Started, to indicate when to start measuring time.
            - Suspended, to pause measuring running time.
            - Resumed, to continue tracking running time after the Timer has been suspended.
            - Restarted, which will cancel previous measurements and start the Timer at zero.

        See Also:
            `ScriptForge.Timer service <https://tinyurl.com/2p84523z>`_
        """
        # region Methods
        @classmethod
        def ReviewServiceArgs(cls, start: bool = ...) -> Tuple[bool]:
            """
            Transform positional and keyword arguments into positional only
            """
        def Continue(self) -> bool:
            """
            Continue a suspended timer.

            Returns:
                bool: True if successful, False if the timer is not suspended
            """
        def Restart(self) -> bool:
            """
            Terminate the timer and restart a new clean timer.

            Returns:
                bool: True if successful, False if the timer is inactive.
            """
        def Start(self) -> bool:
            """
            Start a new clean timer.

            Returns:
                bool: True if successful, False if the timer is already started.
            """
        def Suspend(self) -> bool:
            """
            Suspend a running timer.

            Returns:
                bool: True if successful, False if the timer is not started or already suspended.
            """
        def Terminate(self) -> bool:
            """
            Terminate a running timer.

            Returns:
                bool: True if successful, False if the timer is neither started nor suspended.
            """
        # endregion Methods
        
        # region Properties
        @property
        def Duration(self) -> float:
            """
            Gets the actual running time elapsed since start or between start and stop (does not consider suspended time).
            """
        @property
        def IsStarted(self) -> bool:
            """
            Gets if the timer is started.
            
            True when timer is started or suspended.
            """
        @property
        def IsSuspended(self) -> bool:
            """
            Gets if the timer is suspended.
            
            True when timer is started and suspended.
            """
        @property
        def SuspendDuration(self) -> float:
            """
            Gets the actual time elapsed while suspended since start or between start and stop.
            """
        @property
        def TotalDuration(self) -> float:
            """
            Gets the actual time elapsed since start or between start and stop
            (including suspensions and running time).
            """
        # endregion Properties
    # endregion SF_Timer CLASS
    
    # region SF_UI CLASS
    class SF_UI(SFServices, metaclass=_Singleton):
        """
        Singleton class for the identification and the manipulation of the
        different windows composing the whole LibreOffice application:
            - Windows selection
            - Windows moving and resizing
            - Statusbar settings
            - Creation of new windows
            - Access to the underlying "documents"

        See Also:
            `ScriptForge.UI service <https://tinyurl.com/sez3tpve>`_
        """
        # region CONST
        MACROEXECALWAYS: Literal[2]
        MACROEXECNEVER: Literal[1]
        MACROEXECNORMAL: Literal[0]
        BASEDOCUMENT: Literal["Base"]
        CALCDOCUMENT: Literal["Calc"]
        DRAWDOCUMENT: Literal["Draw"]
        IMPRESSDOCUMENT: Literal["Impress"]
        MATHDOCUMENT: Literal["Math"]
        WRITERDOCUMENT: Literal["Writer"]
        # endregion CONST

        # region Properties
        @property
        def ActiveWindow(self) -> str:
            """
            Gets a valid WindowName for the currently active window.
            When "" is returned, the window could not be identified.
            """
        activeWindow, activewindow = ActiveWindow, ActiveWindow
        # endregion Properties

        # region Methods
        def Activate(self, windowname: str = ...) -> bool:
            """
            Make the specified window active.

            Args:
                windowname (str, optional): see definitions. Defaults to ''.

            Returns:
                bool: True if the given window is found and can be activated.
                There is no change in the actual user interface if no window matches the selection.
            """
        def CreateBaseDocument(
            self,
            filename: str,
            embeddeddatabase: str = ...,
            registrationname: str = ...,
            calcfilename: str = ...,
        ) -> SFDocuments.SF_Base:
            """
            Create a new LibreOffice Base document embedding an empty database of the given type.

            Args:
                filename (str): Identifies the file to create. It must follow the SF_FileSystem.FileNaming notation.
                    If the file already exists, it is overwritten without warning.
                embeddeddatabase (str, optional): either HSQLDB or FIREBIRD or CALC. Defaults to HSQLDB.
                registrationname (str, optional): the name used to store the new database in the databases register.
                    If "" (default), no registration takes place. If the name already exists it is overwritten without warning.
                    Defaults to ''.
                calcfilename (str, optional): only when EmbedddedDatabase = "CALC", the name of the file containing the tables as Calc sheets.
                    The name of the file must be given in SF_FileSystem.FileNaming notation. The file must exist.
                    Defaults to ''.

            Returns:
                SFDocuments.SF_Base: A SFDocuments.SF_Document object or one of its subclasses
            """
        def CreateDocument(
            self, documenttype: str = ..., templatefile: str = ..., hidden: bool = ...
        ) -> SFDocuments.SF_Base:
            """
            Create a new LibreOffice document of a given type or based on a given template.

            Args:
                documenttype (str, optional): Calc, Writer, etc. If absent, a TemplateFile must be given. Defaults to ''.
                templatefile (str, optional): the full FileName of the template to build the new document on.
                    If the file does not exist, the argument is ignored.
                    The "FileSystem" service provides the TemplatesFolder and UserTemplatesFolder
                    properties to help to build the argument. Defaults to ''.
                hidden (bool, optional): if True, open in the background. Defaults to False.
                    To use with caution: activation or closure can only happen programmatically.

            Returns:
                SFDocuments.SF_Base: A SFDocuments.SF_Document object or one of its subclasses
            """
        def Documents(self) -> Tuple[str, ...]:
            """
            Returns the list of the currently open documents. Special windows are ignored.

            Returns:
                Tuple[str, ...]: A tuple of filenames (in SF_FileSystem.FileNaming notation)
                    or of window titles for unsaved documents.
            """
        def GetDocument(
            self, windowname: Union[str, XComponent, DatabaseDocument] = ...
        ) -> Union[SFDocuments.SF_Base, SFDocuments.SF_Calc, SFDocuments.SF_Chart, SFDocuments.SF_Document, SFDocuments.SF_Form, SFDocuments.SF_Writer]:
            """
            Returns a Document object referring to the active window or the given window.

            Args:
                windowname (str | XComponent | DatabaseDocument], optional): when a string, see definitions. If absent the active window is considered.
                    When an object, must be a UNO object of types XComponent or DatabaseDocument.
                    Defaults to ''.

            Returns:
                SFDocuments.SF_Document: document
            """
        def Maximize(self, windowname: Union[str, XWindow] = ...) -> None:
            """
            Maximizes the active window or the given window.

            Args:
                windowname (str | XWindow, optional): see definitions.
                    If absent the active window is considered. Defaults to ''.
            """
        def Minimize(self, windowname: Union[str, XWindow] = ...) -> None:
            """
            Minimizes the current window or the given window

            Args:
                windowname (str | XWindow, optional): see definitions.
                    If absent the active window is considered. Defaults to ''.
            """
        def OpenBaseDocument(
            self,
            filename: str = ...,
            registrationname: str = ...,
            macroexecution: int = ...,
        ) -> SFDocuments.SF_Base:
            """
            Open an existing LibreOffice Base document and return a SFDocuments.Document object.

            Args:
                filename (str, optional): Identifies the file to open. It must follow the SF_FileSystem.FileNaming notation. Defaults to ''.
                registrationname (str, optional): the name of a registered database.
                    It is ignored if FileName <> "". Defaults to ''.
                macroexecution (int, optional): one of the MACROEXECxxx constants. Defaults to MACROEXECNORMAL.

            Returns:
                SFDocuments.SF_Base: A SFDocuments.SF_Base object.
            """
        def OpenDocument(
            self,
            filename: str,
            password: str = ...,
            readonly: bool = ...,
            hidden: bool = ...,
            macroexecution: int = ...,
            filtername: str = ...,
            filteroptions: str = ...,
        ) -> SFDocuments.SF_Document:
            """
            Open an existing LibreOffice document with the given options.

            Args:
                filename (str): Identifies the file to open. It must follow the SF_FileSystem.FileNaming notation.
                password (str, optional): To use when the document is protected.If wrong or absent while the document is protected,
                    the user will be prompted to enter a password. Defaults to ''.
                readonly (bool, optional): Defaults to False.
                hidden (bool, optional): If True, open in the background. Defaults to False.
                    To use with caution: activation or closure can only happen programmatically
                macroexecution (int, optional): One of the MACROEXECxxx constants. Defaults to MACROEXECNORMAL.
                filtername (str, optional): _description_. Defaults to ''.
                filteroptions (str, optional): the name of a filter that should be used for loading the document.If present,
                    the filter must exist. Defaults to ''.

            Returns:
                SFDocuments.SF_Document: A SFDocuments.SF_Document object or one of its subclasses.
                None if the opening failed, including when due to a user decision.
            """
        def Resize(
            self, left: int = ..., top: int = ..., width: int = ..., height: int = ...
        ) -> None:
            """
            Resizes and/or moves the active window. Negative arguments are ignored.
            If the window was minimized or without arguments, it is restored.

            Args:
                left (int, optional): Distances from top of screen. Defaults to -1.
                top (int, optional): Distances from left edge of screen. Defaults to -1.
                width (int, optional): Width of window. Defaults to -1.
                height (int, optional): Height of window. Defaults to -1.
            """
        def SetStatusbar(self, text: str = ..., percentage: int = ...) -> None:
            """
            Display a text and a progressbar in the status bar of the active window.
            Any subsequent calls in the same macro run refer to the same status bar of the same window,
            even if the window is not active anymore.
            A call without arguments resets the status bar to its normal state.

            Args:
                text (str, optional): The optional text to be displayed before the progress bar. Defaults to ''.
                percentage (int, optional): The optional degree of progress between 0 and 100. Defaults to -1.
            """
        def ShowProgressBar(
            self, title: str = ..., text: str = ..., percentage: int = ...
        ) -> None:
            """
            Display a non-modal dialog box. Specify its title, an explicatory text and the progress on a progressbar.
            A call without arguments erases the progress bar dialog.
            The box will anyway vanish at the end of the macro run.

            Args:
                title (str, optional): The title appearing on top of the dialog box (Default = "ScriptForge"). Defaults to ''.
                text (str, optional): The optional text to be displayed above the progress bar. Defaults to ''.
                percentage (int, optional): the degree of progress between 0 and 100. Defaults to -1.
            """
        def WindowExists(self, windowname: Union[str, XWindow]) -> bool:
            """
            Returns True if the specified window exists

            Args:
                windowname (str | XWindow): window

            Returns:
                bool: True if the given window is found.
            """
        # endregion Methods
    # endregion SF_UI CLASS
# endregion SFScriptForge CLASS    (alias of ScriptForge Basic library)

# region SFDatabases CLASS    (alias of SFDatabases Basic library)
class SFDatabases:
    """
    The SFDatabases class manages databases embedded in or connected to Base documents
    """

    # region SF_Database CLASS
    class SF_Database(SFServices):
        """
        Each instance of the current class represents a single database, with essentially its tables, queries
        and data
        The exchanges with the database are done in SQL only.
        To make them more readable, use optionally square brackets to surround table/query/field names
        instead of the (RDBMS-dependent) normal surrounding character.
        SQL statements may be run in direct or indirect mode. In direct mode the statement is transferred literally
        without syntax checking nor review to the database engine.

        See Also:
            `SFDatabases.Database service <https://tinyurl.com/yd9y6xa7>`_
        """
        # region Methods
        @classmethod
        def ReviewServiceArgs(
            cls,
            filename: str = ...,
            registrationname: str = ...,
            readonly: bool = ...,
            user: str = ...,
            password: str = ...,
        ) -> Tuple[str, str, bool, str, str]:
            """
            Transform positional and keyword arguments into positional only
            """
        def CloseDatabase(self) -> None:
            """
            Close the current database connection

            See Also:
                `SF_Database Help CloseDatabase <https://tinyurl.com/yd9y6xa7#CloseDatabase>`_
            """
        def DAvg(
            self, expression: str, tablename: str, criteria: str = ...
        ) -> Union[float, None]:
            """
            Compute the aggregate function AVG() on a  field or expression belonging to a table
            filtered by a WHERE-clause.

            Args:
                expression (str): An SQL expression.
                tablename (str): The name of a table.
                criteria (str, optional): An optional WHERE clause without the word WHERE. Defaults to ''.

            Returns:
                Union[int, float, None]: result

            See Also:
                `SF_Database Help DFunctions <https://tinyurl.com/yd9y6xa7#DFunctions>`_
            """
        def DCount(
            self, expression: str, tablename: str, criteria: str = ...
        ) -> Union[int, None]:
            """
            Compute the aggregate function COUNT() on a  field or expression belonging to a table
            filtered by a WHERE-clause.

            Args:
                expression (str): An SQL expression.
                tablename (str): The name of a table.
                criteria (str, optional): An optional WHERE clause without the word WHERE. Defaults to ''.

            Returns:
                Union[int, None]: result

            See Also:
                `SF_Database Help DFunctions <https://tinyurl.com/yd9y6xa7#DFunctions>`_
            """
        def DLookup(
            self,
            expression: str,
            tablename: str,
            criteria: str = ...,
            orderclause: str = ...,
        ) -> Any:
            """
            Compute the aggregate function Lookup() on a  field or expression belonging to a table
            filtered by a WHERE-clause.

            Args:
                expression (str): An SQL expression.
                tablename (str): The name of a table.
                criteria (str, optional): An optional WHERE clause without the word WHERE. Defaults to ''.
                orderclause (str, optional): An optional order clause incl. "DESC" if relevant. Defaults to ''.

            Returns:
                Any: result

            See Also:
                `SF_Database Help DLookup <https://tinyurl.com/yd9y6xa7#DLookup>`_
            """
        def DMax(
            self, expression: str, tablename: str, criteria: str = ...
        ) -> Union[float, None]:
            """
            Compute the aggregate function MAX() on a  field or expression belonging to a table
            filtered by a WHERE-clause.

            Args:
                expression (str): An SQL expression.
                tablename (str): The name of a table.
                criteria (str, optional): An optional WHERE clause without the word WHERE. Defaults to ''.
            Returns:
                Union[int, float, None]: result

            See Also:
                `SF_Database Help DFunctions <https://tinyurl.com/yd9y6xa7#DFunctions>`_
            """
        def DMin(
            self, expression: str, tablename: str, criteria: str = ...
        ) -> Union[float, None]:
            """
            Compute the aggregate function MIN() on a  field or expression belonging to a table
            filtered by a WHERE-clause.

            Args:
                expression (str): An SQL expression.
                tablename (str): The name of a table.
                criteria (str, optional): An optional WHERE clause without the word WHERE. Defaults to ''.
            Returns:
                Union[int, float, None]: result

            See Also:
                `SF_Database Help DFunctions <https://tinyurl.com/yd9y6xa7#DFunctions>`_
            """
        def DSum(
            self, expression: str, tablename: str, criteria: str = ...
        ) -> Union[float, None]:
            """
            Compute the aggregate function Sum() on a  field or expression belonging to a table
            filtered by a WHERE-clause.

            Args:
                expression (str): An SQL expression.
                tablename (str): The name of a table.
                criteria (str, optional): An optional WHERE clause without the word WHERE. Defaults to ''.
            Returns:
                Union[int, float, None]: result

            See Also:
                `SF_Database Help DFunctions <https://tinyurl.com/yd9y6xa7#DFunctions>`_
            """
        def GetRows(
            self,
            sqlcommand: str,
            directsql: bool = ...,
            header: bool = ...,
            maxrows: int = ...,
        ) -> list:
            """
            Return the content of a table, a query or a SELECT SQL statement as a list

            Args:
                sqlcommand (str): A table name, a query name or a SELECT SQL statement.
                directsql (bool, optional): When True, no syntax conversion is done by LO.
                    Ignored when SQLCommand is a table or a query name. Defaults to False.
                header (bool, optional): When True, a header row is inserted on the top of the array with the column names. Defaults to False.
                maxrows (int, optional): The maximum number of returned rows. If absent, all records are returned. Defaults to 0.

            Returns:
                list: a 2D list[row, column], even if only 1 column and/or 1 record.
                    An empty list if no records returned

            See Also:
                `SF_Database Help GetRows <https://tinyurl.com/yd9y6xa7#GetRows>`_
            """
        def RunSql(self, sqlcommand: str, directsql: bool = ...) -> Any:
            """
            Execute an action query (table creation, record insertion, ...) or SQL statement on the current database.

            Args:
                sqlcommand (str): a query name or an SQL statement
                directsql (bool, optional): when True, no syntax conversion is done by LO.
                    Ignored when SQLCommand is a query name. Defaults to False.

            Returns:
                Any: result

            See Also:
                `SF_Database Help RunSql <https://tinyurl.com/yd9y6xa7#RunSql>`_
            """
        # endregion Methods

        # region Properties
        @property
        def Queries(self) -> Tuple[str, ...]:
            """Gets the list of stored queries."""
        @property
        def Tables(self) -> Tuple[str, ...]:
            """Gets the list of stored tables."""
        @property
        def XConnection(self) -> UNOXConnection:
            """
            Gets the UNO object representing the current database connection.
            """
        @property
        def XMetaData(self) -> XDatabaseMetaData:
            """
            Gets the UNO object representing the metadata describing the database system attributes.
            """
        # endregion Properties

    # endregion SF_Database CLASS
# endregion SFDatabases CLASS    (alias of SFDatabases Basic library)

# region SFDialogs CLASS    (alias of SFDialogs Basic library)
class SFDialogs:
    """
    The SFDialogs class manages dialogs defined with the Basic IDE
    """

    # region SF_Dialog CLASS
    class SF_Dialog(SFServices):
        """
        Each instance of the current class represents a single dialog box displayed to the user.
        The dialog box must have been designed and defined with the Basic IDE previously.
        From a Python script, a dialog box can be displayed in modal or in non-modal modes.

        In modal mode, the box is displayed and the execution of the macro process is suspended
        until one of the OK or Cancel buttons is pressed. In the meantime, other user actions
        executed on the box can trigger specific actions.

        In non-modal mode, the floating dialog remains displayed until the dialog is terminated
        by code (Terminate()) or until the LibreOffice application stops.

        See Also:
            `SFDialogs.Dialog service <https://tinyurl.com/yckkehha>`_
        """
        # region CONST
        # Class constants used together with the Execute() method
        OKBUTTON: Literal[1]
        CANCELBUTTON: Literal[0]
        # endregion CONST

        # region Methods
        @classmethod
        def ReviewServiceArgs(
            cls, container: str = ..., library: str = ..., dialogname: str = ...
        ) -> Tuple[str, str, str, XComponentContext]:
            """
                Transform positional and keyword arguments into positional only
                Add the XComponentContext as last argument
                """
        def Activate(self) -> bool:
            """
            Set the focus on the current dialog instance
            Probably called from after an event occurrence or to focus on a non-modal dialog.

            Returns:
                bool: True if focusing is successful

            See Also:
                `SF_Dialog Help Activate <https://tinyurl.com/yckkehha#Activate>`_
            """
        def Controls(
            self, controlname: str = ...
        ) -> Union[Tuple[str, ...], SFDialogs.SF_DialogControl]:
            """
            Retrun the list of the controls contained in the dialog or
            a dialog control object based on its name.

            Args:
                controlname (str, optional): A valid control name as a case-sensitive string. If absent the list is returned. Defaults to ''.

            Returns:
                Union[Tuple[str, ...], SFDialogs.SF_DialogControl]: A zero-base array of strings if ControlName is absent.
                    An instance of the SFDialogs.SF_DialogControl class if ControlName exists

            See Also:
                `SF_Dialog Help Controls <https://tinyurl.com/yckkehha#Controls>`_
            """
        def EndExecute(self, returnvalue: int) -> None:
            """
            Ends the display of a modal dialog and gives back the argument
            as return value for the current Execute() action.
            
            EndExecute is usually contained in the processing of a macro
            triggered by a dialog or control event.

            Args:
                returnvalue (int): The value passed to the running Execute() method.

            See Also:
                `SF_Dialog Help EndExecute <https://tinyurl.com/yckkehha#EndExecute>`_
            """
        def Execute(self, modal: bool = ...) -> int:
            """
            Display the dialog and wait for its termination by the user.

            Args:
                modal (bool, optional): False when non-modal dialog. Defaults to True.

            Returns:
                int: 0 = Cancel button pressed. 1 = OK button pressed.
                    Otherwise: the dialog stopped with an EndExecute statement executed from a dialog or control event.

            See Also:
                `SF_Dialog Help Execute <https://tinyurl.com/yckkehha#Execute>`_
            """
        def GetTextsFromL10N(self, l10n: SFScriptForge.SF_L10N) -> bool:
            """
            Replace all fixed text strings of a dialog by their localized version.
            
            Replaced texts are:
                - the title of the dialog
                - the caption associated with next control types: Button, CheckBox, FixedLine, FixedText, GroupBox and RadioButton
                - the content of list- and comboboxes
                - the tip- or helptext displayed when the mouse is hovering the control
            
            The current method has a twin method ScriptForge.SF_L10N.AddTextsFromDialog.
            The current method is probably run before the Execute() method.

            Args:
                l10n (SFScriptForge.SF_L10N): A "L10N" service instance created with CreateScriptService("L10N").

            Returns:
                bool: True when successful.

            See Also:
                `SF_Dialog Help GetTextsFromL10N <https://tinyurl.com/yckkehha#GetTextsFromL10N>`_
            """
        def Terminate(self) -> bool:
            """
            Terminate the dialog service for the current dialog instance.
            After termination any action on the current instance will be ignored

            Returns:
                bool: True if termination is successful.

            See Also:
                `SF_Dialog Help Terminate <https://tinyurl.com/yckkehha#Terminate>`_
            """
        # endregion Methods
        
        # region Properties
        @property
        def Caption(self) -> str:
            """Gets/Sets the title of the dialog."""
        @property
        def Height(self) -> int:
            """Gets/Sets the height of the dialog box."""
        @property
        def Modal(self) -> int:
            """Gets if the dialog box is currently in execution in modal mode."""
        @property
        def Name(self) -> str:
            """Gets the name of the dialog"""
        @property
        def OnFocusGained(self) -> str:
            """
            Gets When receiving focus

            Returns:
                str: A URI string with the reference to the script triggered by the event.

            See Also:
                `scripting framework URI specification <https://tinyurl.com/mr4y8k2s>`_.
            """

        @property
        def OnFocusLost(self) -> str:
            """
            Gets When losing focus

            Returns:
                str: A URI string with the reference to the script triggered by the event.

            See Also:
                `scripting framework URI specification <https://tinyurl.com/mr4y8k2s>`_.
            """

        @property
        def OnKeyPressed(self) -> str:
            """
            Gets Key pressed

            Returns:
                str: A URI string with the reference to the script triggered by the event.

            See Also:
                `scripting framework URI specification <https://tinyurl.com/mr4y8k2s>`_.
            """

        @property
        def OnKeyReleased(self) -> str:
            """
            Gets Key released

            Returns:
                str: A URI string with the reference to the script triggered by the event.

            See Also:
                `scripting framework URI specification <https://tinyurl.com/mr4y8k2s>`_.
            """
        @property
        def OnMouseDragged(self) -> str:
            """
            Gets Mouse moved while key presses

            Returns:
                str: A URI string with the reference to the script triggered by the event.

            See Also:
                `scripting framework URI specification <https://tinyurl.com/mr4y8k2s>`_.
            """
        @property
        def OnMouseEntered(self) -> str:
            """
            Gets Mouse inside

            Returns:
                str: A URI string with the reference to the script triggered by the event.

            See Also:
                `scripting framework URI specification <https://tinyurl.com/mr4y8k2s>`_.
            """
        @property
        def OnMouseExited(self) -> str:
            """
            Gets Mouse outside

            Returns:
                str: A URI string with the reference to the script triggered by the event.

            See Also:
                `scripting framework URI specification <https://tinyurl.com/mr4y8k2s>`_.
            """
        @property
        def OnMouseMoved(self) -> str:
            """
            Gets Mouse moved

            Returns:
                str: A URI string with the reference to the script triggered by the event.

            See Also:
                `scripting framework URI specification <https://tinyurl.com/mr4y8k2s>`_.
            """
        @property
        def OnMousePressed(self) -> str:
            """
            Gets Mouse button pressed

            Returns:
                str: A URI string with the reference to the script triggered by the event.

            See Also:
                `scripting framework URI specification <https://tinyurl.com/mr4y8k2s>`_.
            """
        @property
        def OnMouseReleased(self) -> str:
            """
            Gets Mouse button released

            Returns:
                str: A URI string with the reference to the script triggered by the event.

            See Also:
                `scripting framework URI specification <https://tinyurl.com/mr4y8k2s>`_.
            """
        @property
        def Page(self) -> int:
            """
            Gets/Sets A dialog may have several pages that can be traversed by 
            the user step by step. The Page property of the Dialog object defines
            which page of the dialog is active.
            """
        @property
        def Visible(self) -> int:
            """
            Gets/Sets if the dialog box is visible on the desktop.
            By default it is not visible until the Execute() method is run
            and visible afterwards.
            """
        @property
        def XDialogModel(self) -> XControlModel:
            """
            Gets the UNO object representing the dialog model.
            
            Refer to XControlModel and UnoControlDialogModel in Application Programming
            Interface (API) documentation for detailed information.
            """
        @property
        def XDialogView(self) -> XControl:
            """
            Gets the UNO object representing the dialog view.
            
            Refer to XControl and UnoControlDialog in Application Programming Interface
            (API) documentation for detailed information.
            """
        @property
        def Width(self) -> int:
            """Gets/Sets the width of the dialog box."""
        # endregion Properties
    # endregion SF_Dialog CLASS
    
    # region SF_DialogControl CLASS
    class SF_DialogControl(SFServices):
        """
        Each instance of the current class represents a single control within a dialog box.
        The focus is clearly set on getting and setting the values displayed by the controls of the dialog box,
        not on their formatting.
        A special attention is given to controls with type TreeControl.

        See Also:
            `SFDialogs.DialogControl service <https://tinyurl.com/yb27tk36>`_
        """

        # region Methods
        def AddSubNode(
            self, parentnode: XMutableTreeNode, displayvalue: str, datavalue: str = ...,
        ) -> XMutableTreeNode:
            """
            Return a new node of the tree control subordinate to a parent node.

            Args:
                parentnode (XMutableTreeNode): A node UNO object, of type com.sun.star.awt.tree.XMutableTreeNode
                displayvalue (str): The text appearing in the control box
                datavalue (str, optional): Any value associated with the new node. Defaults to ScriptForge.cstSymEmpty.

            Returns:
                XMutableTreeNode: The new node UNO object: com.sun.star.awt.tree.XMutableTreeNode

            See Also:
                `SF_DialogControl Help AddSubNode <https://tinyurl.com/yb27tk36#AddSubNode>`_
            """
        def AddSubTree(
            self,
            parentnode: XMutableTreeNode,
            flattree: list,
            withdatavalue: bool = ...,
        ):
            """
            Return True when a subtree, subordinate to a parent node, could be inserted successfully in a tree control.
            If the parent node had already child nodes before calling this method, the child nodes are erased.
            

            Args:
                parentnode (XMutableTreeNode): A node UNO object, of type com.sun.star.awt.tree.XMutableTreeNode
                flattree (list): a 2D array sorted on the columns containing the DisplayValues.
                withdatavalue (bool, optional): When False (default), every column of FlatTree contains the text to be displayed in the tree control.
                    When True, the texts to be displayed (DisplayValue) are in columns 0, 2, 4, ...
                    while the DataValues are in columns 1, 3, 5, ...
                    Defaults to False.
            
            Notes:
                FlatTree:
                
                .. code::
                
					Flat tree		>>>>		Resulting subtree
					A1	B1	C1					|__	A1
					A1	B1	C2						|__	B1
					A1	B2	C3							|__	C1
					A2	B3	C4							|__	C2
					A2	B3	C5						|__	B2
					A3	B4	C6							|__	C3
													|__	A2
													|__	B3
														|__	C4
														|__	C5
												|__	A3
													|__	B4
														|__	C6

                Typically, such an array can be issued by the GetRows method applied on the SFDatabases.Database service
                when an array item containing the text to be displayed is = "" or is empty/null,
                no new subnode is created and the remainder of the row is skipped.
                When AddSubTree() is called from a Python script, FlatTree may be an list of list.

            See Also:
                `SF_DialogControl Help AddSubTree <https://tinyurl.com/yb27tk36#AddSubTree>`_
            """
        def CreateRoot(
            self, displayvalue: str, datavalue: str = ...
        ) -> XMutableTreeNode:
            """
            Return a new root node of the tree control. The new tree root is inserted below pre-existing root nodes.

            Args:
                displayvalue (str): The text appearing in the control box.
                datavalue (str, optional): Any value associated with the root node. Defaults to ScriptForge.cstSymEmpty.

            Returns:
                XMutableTreeNode: The new root node as a UNO object of type com.sun.star.awt.tree.XMutableTreeNode

            See Also:
                `SF_DialogControl Help CreateRoot <https://tinyurl.com/yb27tk36#CreateRoot>`_
            """
        def FindNode(
            self,
            displayvalue: str,
            datavalue: str = ScriptForge.cstSymEmpty,
            casesensitive: bool = False,
        ) -> Union[XMutableTreeNode, None]:
            """
            Traverses the tree and find recursively, starting from the root, a node meeting some criteria.
            Either (1 match is enough) having its DisplayValue like Dis playValue or
            having its DataValue = DataValue.
            
            Comparisons may be or not case-sensitive.
            
            The first matching occurrence is returned.

            Args:
                displayvalue (str): The pattern to be matched
                datavalue (str, optional): A string, a numeric value or a date or Empty (if not applicable). Defaults to ScriptForge.cstSymEmpty.
                casesensitive (bool, optional): a string, a numeric value or a date or Empty (if not applicable). Defaults to False.

            Returns:
                Union[XMutableTreeNode, None]: The found node of type com.sun.star.awt.tree.XMutableTreeNode or None if not found.

            See Also:
                `SF_DialogControl Help FindNode <https://tinyurl.com/yb27tk36#FindNode>`_
            """
        def SetFocus(self) -> bool:
            """
            Set the focus on the current Control instance.
            Probably called from after an event occurrence.

            Returns:
                bool: True if focusing is successful

            See Also:
                `SF_DialogControl Help SetFocus <https://tinyurl.com/yb27tk36#SetFocus>`_
            """
        def SetTableData(
            self, dataarray: tuple, widths: Tuple[int, ...] = ..., alignments: str = ...
        ) -> bool:
            """
            Fill a table control with the given data. Preexisting data is erased.
            
            The Basic IDE allows to define if the control has a row and/or a column header.
            When it is the case, the array in argument should contain those headers resp. in the first
            column and/or in the first row.
            
            A column in the control shall be sortable when the data (headers excluded) in that column
            is homogeneously filled either with numbers or with strings.
            
            Columns containing strings will be left-aligned, those with numbers will be right-aligned,

            Args:
                dataarray (tuple): the set of data to display in the table control, including optional column/row headers
                    Is a 2D array in Basic, is a tuple of tuples when called from Python
                widths (Tuple[int, ...], optional): tuple containing the relative widths of each column.
                    In other words, widths = Array(1, 2) means that the second column is twice as wide as
                    the first one. If the number of values in the array is smaller than the number of
                    columns in the table, then the last value in the array is used to define the width
                    of the remaining columns.
                alignments (str, optional): the column's horizontal alignment as a string with length = number of columns.
                    Possible characters are: L(EFT), C(ENTER), R(IGHT) or space (default behaviour)
                    Defaults to ''.

            Returns:
                bool: True when successful.

            See Also:
                `SF_DialogControl Help SetTableData <https://tinyurl.com/yb27tk36#SetTableData>`_
            """
        def WriteLine(self, line: str = ...) -> bool:
            """
            Add a new line to a multiline TextField control.

            Args:
                line (str, optional): the line to insert at the end of the text box
                    a newline character will be inserted before the line, if relevant. Defaults to ''.

            Returns:
                bool: True if insertion is successful

            See Also:
                `SF_DialogControl Help WriteLine <https://tinyurl.com/yb27tk36#WriteLine>`_
            """
        # endregion Methods

        # region Properties
        # Root related properties do not start with X and, nevertheless, return a UNO object
        @property
        def Cancel(self) -> bool:
            """
            Gets/Sets if a command button has or not the behaviour of a Cancel button.
            
            Applicable Controls:
                Button
            """
        @property
        def Caption(self) -> str:
            """
            Gets/Sets the text associated with the control.

            Applicable Controls:
                - Button
                - CheckBox
                - FixedLine
                - FixedText
                - GroupBox
                - RadioButton
            """
        @property
        def ControlType(self) -> str:
            """
            Gets/Sets the text associated with the control.

            Applicable Controls:
                All
            """
        @property
        def CurrentNode(self) -> XMutableTreeNode:
            """
            The CurrentNode property returns the currently selected node.
            It returns Empty when there is no node selected.
            When there are several selections, it returns the topmost node among the selected ones.
            
            Applicable Controls:
                TreeControl
            """
        @property
        def Default(self) -> bool:
            """
            Gets/Sets if a command button is the default (OK) button.
            
            Applicable Controls:
                Button
            """
        @property
        def Enabled(self) -> bool:
            """
            Gets/Sets if the control is accessible with the cursor.
            
            Applicable Controls:
                All
            """
        @property
        def Format(self) -> str:
            """
            Gets/Sets the format used to display dates and times.
            
            For dates: "Standard (short)", "Standard (short YY)", "Standard (short YYYY)",
            "Standard (long)", "DD/MM/YY", "MM/DD/YY", "YY/MM/DD", "DD/MM/YYYY", "MM/DD/YYYY",
            "YYYY/MM/DD", "YY-MM-DD", "YYYY-MM-DD".

            For times: "24h short", "24h long", "12h short", "12h long".

            Applicable Controls:
                - DateField
                - TimeField
                - FormattedField (read-only)
            """
        @property
        def ListCount(self) -> int:
            """
            Gets the number of rows in a ListBox, a ComboBox or a TableControl.
            
            Applicable Controls:
                - ComboBox
                - ListBox
                - TableControl
            """
        @property
        def ListIndex(self) -> int:
            """
            Gets/Sets which item is selected in a ListBox, a ComboBox or a TableControl.
            
            Applicable Controls:
                - ComboBox
                - ListBox
                - TableControl
            """
        @property
        def Locked(self) -> bool:
            """
            Gets/Sets if the control is read-only.
            
            Applicable Controls:
                - ComboBox
                - CurrencyField
                - DateField
                - FileControl
                - FormattedField
                - ListBox
                - NumericField
                - PatternField
                - TextField
                - TimeField
            """
        @property
        def MultiSelect(self) -> bool:
            """
            Gets/Sets if a user can make multiple selections in a listbox.
            
            Applicable Controls:
                ListBox
            """
        @property
        def Name(self) -> str:
            """
            Gets the name of the control.

            Applicable Controls:
                All
            """
        @property
        def OnActionPerformed(self) -> str:
            """
            Gets Execute action

            Returns:
                str: A URI string with the reference to the script triggered by the event.

            See Also:
                `scripting framework URI specification <https://tinyurl.com/mr4y8k2s>`_.
            """
        @property
        def OnAdjustmentValueChanged(self) -> str:
            """
            Gets While adjusting

            Returns:
                str: A URI string with the reference to the script triggered by the event.

            See Also:
                `scripting framework URI specification <https://tinyurl.com/mr4y8k2s>`_.
            """
        @property
        def OnFocusGained(self) -> str:
            """
            Gets When receiving focus

            Returns:
                str: A URI string with the reference to the script triggered by the event.

            See Also:
                `scripting framework URI specification <https://tinyurl.com/mr4y8k2s>`_.
            """
        @property
        def OnFocusLost(self) -> str:
            """
            Gets When losing focus

            Returns:
                str: A URI string with the reference to the script triggered by the event.

            See Also:
                `scripting framework URI specification <https://tinyurl.com/mr4y8k2s>`_.
            """
        @property
        def OnItemStateChanged(self) -> str:
            """
            Gets Item status changed

            Returns:
                str: A URI string with the reference to the script triggered by the event.

            See Also:
                `scripting framework URI specification <https://tinyurl.com/mr4y8k2s>`_.
            """
        @property
        def OnKeyPressed(self) -> str:
            """
            Gets Key pressed

            Returns:
                str: A URI string with the reference to the script triggered by the event.

            See Also:
                `scripting framework URI specification <https://tinyurl.com/mr4y8k2s>`_.
            """
        @property
        def OnKeyReleased(self) -> str:
            """
            Gets Key released

            Returns:
                str: A URI string with the reference to the script triggered by the event.

            See Also:
                `scripting framework URI specification <https://tinyurl.com/mr4y8k2s>`_.
            """
        @property
        def OnMouseDragged(self) -> str:
            """
            Gets Mouse moved while key presses

            Returns:
                str: A URI string with the reference to the script triggered by the event.

            See Also:
                `scripting framework URI specification <https://tinyurl.com/mr4y8k2s>`_.
            """
        @property
        def OnMouseEntered(self) -> str:
            """
            Gets Mouse inside

            Returns:
                str: A URI string with the reference to the script triggered by the event.

            See Also:
                `scripting framework URI specification <https://tinyurl.com/mr4y8k2s>`_.
            """
        @property
        def OnMouseExited(self) -> str:
            """
            Gets Mouse outside

            Returns:
                str: A URI string with the reference to the script triggered by the event.

            See Also:
                `scripting framework URI specification <https://tinyurl.com/mr4y8k2s>`_.
            """
        @property
        def OnMouseMoved(self) -> str:
            """
            Gets Mouse moved

            Returns:
                str: A URI string with the reference to the script triggered by the event.

            See Also:
                `scripting framework URI specification <https://tinyurl.com/mr4y8k2s>`_.
            """
        @property
        def OnMousePressed(self) -> str:
            """
            Gets Mouse button pressed

            Returns:
                str: A URI string with the reference to the script triggered by the event.

            See Also:
                `scripting framework URI specification <https://tinyurl.com/mr4y8k2s>`_.
            """
        @property
        def OnMouseReleased(self) -> str:
            """
            Gets Mouse button released

            Returns:
                str: A URI string with the reference to the script triggered by the event.

            See Also:
                `scripting framework URI specification <https://tinyurl.com/mr4y8k2s>`_.
            """
        @property
        def OnNodeExpanded(self) -> str:
            """
            Gets (Not in Basic IDE) when the expansion button is pressed on a node in a tree control

            Returns:
                str: A URI string with the reference to the script triggered by the event.

            See Also:
                `scripting framework URI specification <https://tinyurl.com/mr4y8k2s>`_.
            """
        @property
        def OnNodeSelected(self) -> str:
            """
            Gets (Not in Basic IDE) when a node in a tree control is selected

            Returns:
                str: A URI string with the reference to the script triggered by the event.

            See Also:
                `scripting framework URI specification <https://tinyurl.com/mr4y8k2s>`_.
            """
        @property
        def OnTextChanged(self) -> str:
            """
            Gets Text modified

            Returns:
                str: A URI string with the reference to the script triggered by the event.

            See Also:
                `scripting framework URI specification <https://tinyurl.com/mr4y8k2s>`_.
            """
        @property
        def Page(self) -> int:
            """
            Gets/Sets the page of the dialog on which the control is visible.
            
            A dialog may have several pages that can be traversed by the user step by step.
            The Page property of the Dialog object defines which page of the dialog is active.

            Applicable Controls:
                All
            """
        @property
        def Parent(self) -> SFDialogs.SF_Dialog:
            """
            Gets the parent SFDialogs.Dialog class object instance.

            Applicable Controls:
                All
            """
        @property
        def Picture(self) -> str:
            """
            Gets/Sets the file name containing a bitmap or other type of graphic to be
            displayed on the specified control. The filename must comply with the
            FileNaming attribute of the ``ScriptForge.FileSystem`` service.

            Applicable Controls:
                * Button
                * ImageControl
            """
        @property
        def RootNode(self) -> XMutableTreeNode:
            """
            Gets the RootNode property returns the last root node of a tree control

            Applicable Controls:
                TreeControl
            """
        @property
        def RowSource(self) -> Tuple[str, ...]:
            """
            Gets/Sets the data contained in a combobox or a listbox.

            Applicable Controls:
                * ComboBox
                * ListBox
            """
        @property
        def Text(self) -> str:
            """
            Gets the text being displayed by the control.

            Applicable Controls:
                - ComboBox
                - FileControl
                - FormattedField
                - PatternField
                - TextField
            """
        @property
        def TipText(self) -> str:
            """
            Gets/Sets the text that appears as a tooltip when you hold the mouse pointer over the control.

            Applicable Controls:
                All
            """
        @property
        def TripleState(self) -> bool:
            """
            Gets/Sets if the control may have the state "don't know".
            
            Applicable Controls:
                CheckBox
            """
        @property
        def Value(self) -> Any:
            """
            Gets/Sets control type.

            See Also:
                `The Value property <https://tinyurl.com/yb27tk36#hd_id81598540704978>`_
            """
        @property
        def Visible(self) -> bool:
            """
            Gets/Sets if the control is hidden or visible.
            
            Applicable Controls:
                All
            """
        @property
        def XControlModel(self) -> XControlModel:
            """
            Gets the UNO object representing the control model.
            
            Applicable Controls:
                All
            """
        @property
        def XControlView(self) -> XControl:
            """
            Gets the UNO object representing the control view.
            
            Applicable Controls:
                All
            """
        @property
        def XTreeDataModel(self) -> XMutableTreeDataModel:
            """
            Gets the UNO object representing the tree control data model.
            
            Applicable Controls:
                TreeControl
            """
        # endregion Properties
        
    # endregion SF_DialogControl CLASS
# endregion SFDialogs CLASS    (alias of SFDialogs Basic library)

# region SFDocuments CLASS    (alias of SFDocuments Basic library)
class SFDocuments:
    """
    The SFDocuments class gathers a number of classes, methods and properties making easy
    managing and manipulating LibreOffice documents
    """

    # region SF_Document CLASS
    class SF_Document(SFServices):
        """
        The methods and properties are generic for all types of documents: they are combined in the
        current SF_Document class
            - saving, closing documents
            - accessing their standard or custom properties

        Specific properties and methods are implemented in the concerned subclass(es) SF_Calc, SF_Base, ...

        See Also:
            `SFDocuments.Document service <https://tinyurl.com/ybujrgjk>`_
        """
        # region methods
        @classmethod
        def ReviewServiceArgs(cls, windowname: str = ...) -> Tuple[str]:
            """
            Transform positional and keyword arguments into positional only
            """
        def Activate(self) -> bool:
            """
            Make the current document active

            Returns:
                bool: True if the document could be activated; Otherwise, there is no change in the actual user interface.

            See Also:
                `SF_Document Help Activate <https://tinyurl.com/ybujrgjk#Activate>`_
            """
        def CloseDocument(self, saveask: bool = ...) -> bool:
            """
            Close the document. Does nothing if the document is already closed
            regardless of how the document was closed, manually or by program.

            Args:
                saveask (bool, optional): If True, the user is invited to confirm or not the writing of the changes on disk.
                    No effect if the document was not modified. Defaults to True.

            Returns:
                bool: False if the user declined to close

            See Also:
                `SF_Document Help CloseDocument <https://tinyurl.com/ybujrgjk#CloseDocument>`_
            """
        def ExportAsPDF(
            self,
            filename: str,
            overwrite: bool = ...,
            pages: str = ...,
            password: str = ...,
            watermark: str = ...,
        ) -> bool:
            """
            Store the document to the given file location in PDF format

            Args:
                filename (str): Identifies the file where to save. It must follow the SF_FileSystem.FileNaming notation
                overwrite (bool, optional): True if the destination file may be overwritten. Defaults to False.
                pages (str, optional): The pages to print as a string, like in the user interface. Example: "1-4;10;15-18". Defaults to ''.
                password (str, optional): Password to open the document. Defaults to ''.
                watermark (str, optional): The text for a watermark to be drawn on every page of the exported PDF file. Defaults to ''.

            Returns:
                bool: False if the document could not be saved

            See Also:
                `SF_Document Help ExportAsPDF <https://tinyurl.com/ybujrgjk#ExportAsPDF>`_
            """
        def PrintOut(self, pages: str = ..., copies: int = ...) -> bool:
            """
            Send the content of the document to the printer.
            The printer might be defined previously by default, by the user or by the SetPrinter() method.

            Args:
                pages (str, optional): The pages to print as a string, like in the user interface. Example: "1-4;10;15-18".Defaults to ''.
                copies (int, optional): The number of copies. Defaults to 1.

            Returns:
                bool: True when successful.

            See Also:
                `SF_Document Help PrintOut <https://tinyurl.com/ybujrgjk#PrintOut>`_
            """
        def RunCommand(self, command: str) -> None:
            """
            Run on the document the given menu command. The command is executed without arguments.

            Args:
                command (str): Case-sensitive. The command itself is not checked.
                    If nothing happens, then the command is probably wrong.
                    A few typical commands - Save, SaveAs, ExportToPDF, SetDocumentProperties,
                    Undo, Copy, Paste, About

            See Also:
                `SF_Document Help RunCommand <https://tinyurl.com/ybujrgjk#RunCommand>`_
            """
        def Save(self) -> bool:
            """
            Store the document to the file location from which it was loaded.
            Ignored if the document was not modified

            Returns:
                bool: False if the document could not be saved.

            See Also:
                `SF_Document Help Save <https://tinyurl.com/ybujrgjk#Save>`_
            """
        def SaveAs(
            self,
            filename: str,
            overwrite: bool = ...,
            password: str = ...,
            filtername: str = ...,
            filteroptions: str = ...,
        ) -> bool:
            """
            Store the document to the given file location.
            The new location becomes the new file name on which simple Save method calls will be applied.

            Args:
                filename (str): Identifies the file where to save. It must follow the SF_FileSystem.FileNaming notation.
                overwrite (bool, optional): True if the destination file may be overwritten. Defaults to False.
                password (str, optional): Use to protect the document. Defaults to ''.
                filtername (str, optional): The name of a filter that should be used for saving the document
                    If present, the filter must exist. Defaults to ''.
                filteroptions (str, optional): An optional string of options associated with the filter. Defaults to ''.

            Returns:
                bool: False if the document could not be saved.

            See Also:
                `SF_Document Help SaveAs <https://tinyurl.com/ybujrgjk#SaveAs>`_
            """
        def SaveCopyAs(
            self,
            filename: str,
            overwrite: bool = ...,
            password: str = ...,
            filtername: str = ...,
            filteroptions: str = ...,
        ) -> bool:
            """
            Store a copy or export the document to the given file location.
            The actual location is unchanged

            Args:
                filename (str): Identifies the file where to save. It must follow the SF_FileSystem.FileNaming notation.
                overwrite (bool, optional): True if the destination file may be overwritten. Defaults to False.
                password (str, optional): Use to protect the document. Defaults to ''.
                filtername (str, optional): the name of a filter that should be used for saving the document.
                    If present, the filter must exist. Defaults to ''.
                filteroptions (str, optional): An optional string of options associated with the filter. Defaults to ''.

            Returns:
                bool: False if the document could not be saved.

            See Also:
                `SF_Document Help SaveCopyAs <https://tinyurl.com/ybujrgjk#SaveCopyAs>`_
            """
        def SetPrinter(
            self, printer: str = ..., orientation: str = ..., paperformat: str = ...
        ) -> bool:
            """
            Define the printer options for the document.

            Args:
                printer (str, optional): the name of the printer queue where to print to.
                    When absent or space, the default printer is set. Defaults to ''.
                orientation (str, optional): either "PORTRAIT" or "LANDSCAPE". Left unchanged when absent. Defaults to ''.
                paperformat (str, optional): Paper format, see note. Left unchanged when absent.

            Note:
                Arg ``paperformat`` accepted values:
                    - A3
                    - A4
                    - A5
                    - B4
                    - B5
                    - LETTER
                    - LEGAL
                    - TABLOID
            Returns:
                bool: True when successful.

            See Also:
                `SF_Document Help SetPrinter <https://tinyurl.com/ybujrgjk#SetPrinter>`_
            """
        # endregion methods
        
        # region Properties
        @property
        def CustomProperties(self) -> SFScriptForge.SF_Dictionary:
            """
            Gets/Sets a ScriptForge.Dictionary object instance.
            
            After update, can be passed again to the property for updating the document.
            Individual items of the dictionary may be either strings, numbers,
            (Basic) dates or com.sun.star.util.Duration items.
            """
        @property
        def Description(self) -> str:
            """
            Gets/Sets the Description property of the document (also known as "Comments")
            """
        @property
        def DocumentProperties (self) -> SFScriptForge.SF_Dictionary:
            """
            Gets a ScriptForge.Dictionary object containing all the entries.
            
            Document statistics are included. Note that they are specific to the type of document.
            As an example, a Calc document includes a "CellCount" entry. Other documents do not.
            """
        @property
        def DocumentType(self) -> str:
            """
            Gets the value with the document type ("Base", "Calc", "Writer", etc)
            """
        @property
        def IsBase(self) -> bool:
            """
            Gets if instance is Base document.
            """
        @property
        def IsCalc(self) -> bool:
            """
            Gets if instance is Calc document.
            """
        @property
        def IsDraw(self) -> bool:
            """
            Gets if instance is Draw document.
            """
        @property
        def IsImpress(self) -> bool:
            """
            Gets if instance is Draw document.
            """
        @property
        def IsMath(self) -> bool:
            """
            Gets if instance is Math document.
            """
        @property
        def IsWriter(self) -> bool:
            """
            Gets if instance is Writer document.
            """
        @property
        def Keywords(self) -> str:
            """
            Gets/Sets the Keywords property of the document. Represented as a comma-separated list of keywords
            """
        @property
        def Readonly(self) -> bool:
            """
            Gets if the document is actually in read-only mode.
            """
        @property
        def Subject(self) -> str:
            """
            Gets/Sets the Subject property of the document.
            """
        @property
        def Title(self) -> str:
            """
            Gets/Sets the Title property of the document.
            """
        @property
        def XComponent(self) -> XComponent:
            """
            Gets The UNO object com.sun.star.lang.XComponent
            or com.sun.star.comp.dba.OfficeDatabaseDocument representing the document
            """
        # endregion Properties
    # endregion SF_Document CLASS
    
    # region SF_Base CLASS
    class SF_Base(SF_Document, SFServices):
        """
        The SF_Base module is provided mainly to block parent properties that are NOT applicable to Base documents
        In addition, it provides methods to identify form documents and access their internal forms
        (read more elsewhere (the "SFDocuments.Form" service) about this subject)

        See Also:
            `SFDocuments.Base service <https://tinyurl.com/ya4lp2mq>`_
        """
        # region methods
        @classmethod
        def ReviewServiceArgs(cls, windowname: str = "") -> Tuple[str]:
            """
            Transform positional and keyword arguments into positional only
            """
        def CloseDocument(self, saveask: bool = True) -> bool:
            """
            The closure of a Base document requires the closures of
                - the connection => done in the CloseDatabase() method
                - the data source
                - the document itself => done in the superclass

            Args:
                saveask (bool, optional): Ask to save. Defaults to True.

            Returns:
                bool: True if closure is successful.
            """
        def CloseFormDocument(self, formdocument: bool = ...) -> bool:
            """
            Close the given form document.
            If nothing happens if the form document is not open.

            Args:
                formdocument (bool, optional): a valid document form name as a case-sensitive string.

            Returns:
                bool: True if closure is successful.

            See Also:
                `SF_Base Help CloseFormDocument <https://tinyurl.com/ya4lp2mq#CloseFormDocument>`_
            """
        def FormDocuments(self) -> Tuple[str, ...]:
            """
            Return the list of the FormDocuments contained in the Base document

            Returns:
                Tuple[str, ...]: A tuple of strings.
                    Each entry is the full path name of a form document. The path separator is the slash ("/")

            See Also:
                `SF_Base Help FormDocuments <https://tinyurl.com/ya4lp2mq#FormDocuments>`_
            """
        def GetDatabase(
            self, user: str = ..., password: str = ...
        ) -> Union[SFDatabases.SF_Database, None]:
            """
            Returns a Database instance (service = SFDatabases.Database) giving access
            to the execution of SQL commands on the database defined and/or stored in
            the actual Base document

            Args:
                user (str, optional): The Login user of database. Defaults to ''.
                password (str, optional): password of user. Defaults to ''.

            Returns:
                SFDatabases.SF_Database | None: SF_Database or None.

            See Also:
                `SF_Base Help GetDatabase <https://tinyurl.com/ya4lp2mq#GetDatabase>`_
            """
        def IsLoaded(self, formdocument: str) -> bool:
            """
            Return True if the given FormDocument is open for the user

            Args:
                formdocument (str): A valid document form name as a case-sensitive string

            Returns:
                bool: True if the form document is currently open, otherwise False.

            See Also:
                `SF_Base Help IsLoaded <https://tinyurl.com/ya4lp2mq#IsLoaded>`_
            """

        @overload
        def OpenFormDocument(self, formdocument: str) -> bool: ...
        @overload
        def OpenFormDocument(self, formdocument: str, designmode: bool) -> bool:
            """
            Open the FormDocument given by its hierarchical name either in normal or in design mode.
            If the form document is already open, the form document is made active without changing its mode.

            Args:
                formdocument (str): A valid document form name as a case-sensitive string.
                designmode (bool, optional): When True the form document is opened in design mode. Defaults to False.

            Returns:
                bool: True if the form document could be opened; Otherwise False.

            See Also:
                `SF_Base Help OpenFormDocument <https://tinyurl.com/ya4lp2mq#OpenFormDocument>`_
            """
        def PrintOut(
            self, formdocument: str, pages: str = ..., copies: int = ...
        ) -> bool:
            """
            Send the content of the given form document to the printer.
            The printer might be defined previously by default, by the user or by the SetPrinter() method.
            The given form document must be open. It is activated by the method.

            Args:
                formdocument (str): A valid document form name as a case-sensitive string.
                pages (str, optional): The pages to print as a string, like in the user interface. Example: "1-4;10;15-18". Defaults to ''.
                copies (int, optional): The number of copies. Defaults to 1.

            Returns:
                bool: True when successful.

            See Also:
                `SF_Base Help PrintOut <https://tinyurl.com/ya4lp2mq#PrintOut>`_
            """
        def SetPrinter(
            self,
            formdocument: str = ...,
            printer: str = ...,
            orientation: str = ...,
            paperformat: str = ...,
        ) -> bool:
            """
            Define the printer options for a form document. The form document must be open.

            Args:
                formdocument (str, optional): a valid document form name as a case-sensitive string. Defaults to ''.
                printer (str, optional): the name of the printer queue where to print to. When absent or space, the default printer is set. Defaults to ''.
                orientation (str, optional): either ``PORTRAIT`` or ``LANDSCAPE``. Defaults to ''.
                paperformat (str, optional): Paper format, see note. Left unchanged when absent.

            Note:
                Arg ``paperformat`` accepted values:
                    - A3
                    - A4
                    - A5
                    - B4
                    - B5
                    - LETTER
                    - LEGAL
                    - TABLOID

            Returns:
                bool: True when successful.

            See Also:
                `SF_Base Help SetPrinter <https://tinyurl.com/ya4lp2mq#SetPrinter>`_
            """
        # endregion methods
    # endregion SF_Base CLASS
    
    # region SF_Calc CLASS
    class SF_Calc(SF_Document, SFServices):
        """
        The SF_Calc module is focused on :
        - management (copy, insert, move, ...) of sheets within a Calc document
        - exchange of data between Basic data structures and Calc ranges of values

        See Also:
           `SF_Calc Help <https://tinyurl.com/y7jwr7b7>`_
        """
        # region methods
        @classmethod
        def ReviewServiceArgs(cls, windowname: str = ...) -> Tuple[str]:
            """
            Transform positional and keyword arguments into positional only
            """
        # Next functions are implemented in Basic as read-only properties with 1 argument
        def FirstCell(self, rangename: str) -> str:
            """
            Returns the First used cell in a given range or sheet.

            Args:
                rangename (str): SheetName or RangeName as String.
                    When the argument is a sheet it will always return the "sheet.$A$1" cell

            Returns:
                str: Returns the first used cell in a given range or sheet.
            """
        def FirstColumn(self, rangename: str) -> int:
            """
            Returns the leftmost column number in a given range or sheet.

            Args:
                rangename (str): SheetName or RangeName as String.
                    When the argument is a sheet it will always return 1.

            Returns:
                int: Returns the leftmost column number in a given range or sheet.
            """
        def FirstRow(self, rangename: str) -> int:
            """
            Returns the First used column in a given range.

            Args:
                rangename (str): SheetName or RangeName as String.
                    When the argument is a sheet it will always return 1.

            Returns:
                int: 	Returns the First used column in a given range.
            """
        def Height(self, rangename: str) -> int:
            """
            Returns the height in # of rows of the given range.

            Args:
                rangename (str): SheetName or RangeName as String.

            Returns:
                int: Returns the height in # of rows of the given range.
            """
        def LastCell(self, rangename: str) -> str:
            """
            Returns the last used cell in a given sheet or range.

            Args:
                rangename (str): SheetName or RangeName as String.

            Returns:
                str: Returns the last used cell in a given sheet or range.
            """
        def LastColumn(self, rangename: str) -> int:
            """
            Returns the last used column in a given sheet.

            Args:
                rangename (str): SheetName or RangeName as String.

            Returns:
                int: Returns the last used column in a given sheet.
            """
        def LastRow(self, rangename: str) -> int:
            """
            The last used row in a given range or sheet.

            Args:
                rangename (str): SheetName or RangeName as String.

            Returns:
                int: The last used row in a given range or sheet.
            """
        def Range(self, rangename: str) -> SFDocuments.SF_CalcReference:
            """
            Returns a (internal) range object

            Args:
                rangename (range): RangeName as String.

            Returns:
                object: Returns a (internal) range object.
            """
        def Region(self, rangename: str) -> str:
            """
            Returns the address of the smallest area that contains the specified range
            so that the area is surrounded by empty cells or sheet edges.
            This is equivalent to applying the Ctrl + * shortcut to the given range.

            Args:
                rangename (str): RangeName As String

            Returns:
                str: region as string
            """
        def Sheet(self, sheetname: str) -> "SFDocuments.SF_CalcReference":
            """
            Returns a sheet object

            Args:
                sheetname (str): SheetName As String.

            Returns:
                str: Returns a sheet object 
            """
        def SheetName(self, rangename: str) -> str:
            """
                Returns the sheet name part of a rang.

            Args:
                rangename (str): RangeName As String.

            Returns:
                str: Returns the sheet name part of a rang.
            """
        def Width(self, rangename: str) -> int:
            """
            The number of columns (>= 1) in the given range.

            Args:
                rangename (str): RangeName As String.

            Returns:
                int: number of columns
            """
        def XCellRange(self, rangename: str) -> XCellRange:
            """
            A ``com.sun.star.Table.XCellRange`` UNO object.

            Args:
                rangename (str): RangeName As String.

            Returns:
                XCellRange: cell range.
            """
        def XSheetCellCursor(self, rangename: str) -> XSheetCellCursor:
            """
            A ``com.sun.star.sheet.XSheetCellCursor`` UNO object.
            After moving the cursor, the resulting range address can be accessed through
            the AbsoluteName UNO property of the cursor object, which returns a string
            value that can be used as argument for properties and methods of the Calc service.

            Args:
                rangename (str): RangeName As String.

            Returns:
                XSheetCellCursor: sheet cell cursor.
            """
        def XSpreadsheet(self, sheetname: str) -> XSpreadsheet:
            """
            A com.sun.star.sheet.XSpreadsheet UNO object.

            Args:
                sheetname (str): SheetName As String.

            Returns:
                XSpreadsheet: spreadsheet.
            """
        # Usual methods
        def A1Style(
            self,
            row1: int,
            column1: int,
            row2: int = ...,
            column2: int = ...,
            sheetname: str = ...,
        ) -> str:
            """
            Returns a range address as a string based on sheet coordinates, i.e. row and column numbers.

            If only a pair of coordinates is given, then an address to a single cell is returned.
            Additional arguments can specify the bottom-right cell of a rectangular range.
            
            Row and column numbers start at 1.

            Args:
                row1 (int): Specify the row number of the top cell in the range to be considered.
                column1 (int): Specify the column number of the left cell in the range to be considered.
                row2 (int, optional): Specify the row number of the bottom cell in the range to be considered.
                    If this arguments is not provided, or if values smaller than row1 and column1 are given,
                    then the address of the single cell range represented by row1 and column1 is returned. Defaults to 0.
                column2 (int, optional): Specify the column number of the right cell in the range to be considered.
                    If these arguments are not provided, or if values smaller than row1 and column1 are given,
                    then the address of the single cell range represented by row1 and column1 is returned. Defaults to 0.
                sheetname (str, optional): The name of the sheet to be appended to the returned range address.
                    The sheet must exist. The default value is "~" corresponding to the currently active sheet.

            Returns:
                str: Returns a range address as a string.
            
            See Also:
                `SF_Calc Help A1Style <https://tinyurl.com/y7jwr7b7#A1Style>`_
            """
        @overload
        def Activate(self) -> bool:
            """
            Activates the document window.

            Returns:
                bool: True if the document or the sheet could be made active;
                Otherwise, there is no change in the actual user interface.
            
            See Also:
                `SF_Calc Help Activate <https://tinyurl.com/y7jwr7b7#Activate>`_
            """
        @overload
        def Activate(self, sheetname: str) -> bool:
            """
            the given sheet is activated and it becomes the currently selected sheet.

            Args:
                sheetname (str):The name of the sheet to be activated in the document.

            Returns:
                bool: True if the document or the sheet could be made active;
                Otherwise, there is no change in the actual user interface.
            
            See Also:
                `SF_Calc Help Activate <https://tinyurl.com/y7jwr7b7#Activate>`_
            """
        def Charts(
            self, sheetname: str, chartname: str = ...
        ) -> Tuple[str] | "SFDocuments.SF_Chart":
            """
            Returns either the list with the names of all chart objects in a given sheet or a single Chart service instance.
            * If only ``sheetname`` is specified, an zero-based array of strings containing the names of all charts is returned.
            * If a ``chartname`` is provided, than a single object corresponding to the desired chart is returned. The specified chart must exist.
            
            Args:
                sheetname (str): The name of the sheet from which the list of charts is to be retrieved or where the specified chart is located.
                chartname (str, optional): The user-defined name of the chart object to be returned.
                    If the chart does not have a user-defined name, then the internal object name can be used.
                    If this argument is absent, then the list of chart names in the specified sheet is returned. Defaults to ''.

            Returns:
                Tuple[str] | SFDocuments.SF_Chart: Returns either the list with the names of all chart objects in a given sheet or a single Chart service instance.
            
            See Also:
                `SF_Calc Help Charts <https://tinyurl.com/y7jwr7b7#Charts>`_
            """
        def ClearAll(self, range: str) -> None:
            """
            Clears all the contents and formats of the given range.

            Args:
                range (str): The range to be cleared, as a string.
            
            See Also:
                `SF_Calc Help ClearAll <https://tinyurl.com/y7jwr7b7#ClearAll>`_
            """
        def ClearFormats(self, range: str) -> None:
            """
            Clears the formats and styles in the given range.

            Args:
                range (str): The range whose formats and styles are to be cleared, as a string.

            See Also:
                `SF_Calc Help ClearFormats <https://tinyurl.com/y7jwr7b7#ClearFormats>`_
            """
        def ClearValues(self, range: str) -> None:
            """
            Clears the values and formulas in the given range.

            Args:
                range (str): The range whose values and formulas are to be cleared, as a string.
                
            See Also:
                `SF_Calc Help ClearValues <https://tinyurl.com/y7jwr7b7#ClearValues>`_
            """
        @overload
        def CopySheet(self, sheetname: str, newname: str) -> bool:
            """
            Copies a specified sheet at the end of the list of sheets.
            The sheet to be copied may be contained inside any **open** Calc document.

            Args:
                sheetname (str): The name of the sheet to be copied as a string or its reference as an object.
                newname (str): The name of the sheet to insert. The name must not be in use in the document.

            Returns:
                bool: True if the sheet could be copied successfully.

            See Also:
                `SF_Calc Help CopySheet <https://tinyurl.com/y7jwr7b7#CopySheet>`_
            """
        @overload
        def CopySheet(self, sheetname: str, newname: str, beforesheet: int) -> bool:
            """
            Copies a specified sheet before an existing sheet.
            The sheet to be copied may be contained inside any **open** Calc document.

            Args:
                sheetname (str): The name of the sheet to be copied as a string or its reference as an object.
                newname (str): The name of the sheet to insert. The name must not be in use in the document.
                beforesheet (int): The name of the sheet before which to insert the copied sheet.

            Returns:
                bool: True if the sheet could be copied successfully.

            See Also:
                `SF_Calc Help CopySheet <https://tinyurl.com/y7jwr7b7#CopySheet>`_
            """
        @overload
        def CopySheet(self, sheetname: str, newname: str, beforesheet: str) -> bool:
            """
            Copies a specified sheet before an existing sheet or at the end of the list of sheets.
            The sheet to be copied may be contained inside any **open** Calc document.

            Args:
                sheetname (str): The name of the sheet to be copied as a string or its reference as an object.
                newname (str): The name of the sheet to insert. The name must not be in use in the document.
                beforesheet (str): The index (numeric, starting from 1) of the sheet before which to insert the copied sheet.

            Returns:
                bool: True if the sheet could be copied successfully.

            See Also:
                `SF_Calc Help CopySheet <https://tinyurl.com/y7jwr7b7#CopySheet>`_
            """
        def CopySheetFromFile(
            self, filename: str, sheetname: str, newname: str, beforesheet: int = ...
        ) -> bool:
            """
            Copies a specified sheet from a closed Calc document and pastes it before an existing
            sheet or at the end of the list of sheets of the file referred to by a ``Document`` object.
            
            If the file does not exist, an error is raised. If the file is not a valid Calc file,
            a blank sheet is inserted. If the source sheet does not exist in the input file,
            an error message is inserted at the top of the newly pasted sheet.

            Args:
                filename (str): Identifies the file to open. It must follow the SF_FileSystem.FileNaming
                    notation. The file must not be protected with a password.
                sheetname (str): The name of the sheet to be copied as a string.
                newname (str): The name of the copied sheet to be inserted in the document.
                    The name must not be in use in the document.
                beforesheet (int, optional): The name (string) or index (numeric, starting from 1)
                    of the sheet before which to insert the copied sheet. This argument is optional
                    and the default behavior is to add the copied sheet at the last position.

            Returns:
                bool: True if the sheet could be created.

            See Also:
                `SF_Calc Help CopySheetFromFile <https://tinyurl.com/y7jwr7b7#CopySheetFromFile>`_
            """
        def CopyToCell(self, sourcerange: Union[str, Any], destinationcell: str) -> str:
            """
            Copies a specified source range (values, formulas and formats) to a destination range or cell.
            The method reproduces the behaviour of a Copy/Paste operation from a range to a single cell.
            
            It returns a string representing the modified range of cells. The size of the modified area
            is fully determined by the size of the source area.
            
            The source range may belong to another **open** document.

            Args:
                sourcerange (str, Any): The source range as a string when it belongs to the same document or as a reference when it belongs to another open Calc document.
                destinationcell (str): The destination cell where the copied range of cells will be pasted, as a string. If a range is given, only its top-left cell is considered.

            Returns:
                str: A string representing the modified range of cells.
                The modified area depends only on the size of the source area

            See Also:
                `SF_Calc Help CopyToCell <https://tinyurl.com/y7jwr7b7#CopyToCell>`_
            """
        def CopyToRange(self, sourcerange: Union[str, Any], destinationrange: str) -> str:
            """
            Copies downwards and/or rightwards a specified source range (values, formulas and formats)
            to a destination range. The method imitates the behaviour of a Copy/Paste operation from
            a source range to a larger destination range.
            * If the height (or width) of the destination area is > 1 row (or column) then the height (or width)
            of the source must be <= the height (or width) of the destination. Otherwise nothing happens.
            * If the height (or width) of the destination is = 1 then the destination is expanded downwards
            (or rightwards) up to the height (or width) of the source range.
            
            The source range may belong to another **open** document.

            Args:
                sourcerange (str | Any): The source range as a string when it belongs to the same document
                    or as a reference when it belongs to another open Calc document.
                destinationrange (str): The destination of the copied range of cells, as a string.

            Returns:
                str: a string representing the modified range of cells.

            See Also:
                `SF_Calc Help CopyToRange <https://tinyurl.com/y7jwr7b7#CopyToRange>`_
            """
        def CreateChart(
            self,
            chartname: str,
            sheetname: str,
            range: str,
            columnheader: bool = ...,
            rowheader: bool = ...,
        ) -> SFDocuments.SF_Chart:
            """
            Creates a new chart object showing the data in the specified range.
            The returned chart object can be further manipulated using the ``Chart service``.

            Args:
                chartname (str): The user-defined name of the chart to be created. The name must be unique in the same sheet.
                sheetname (str): The name of the sheet where the chart will be placed.
                range (str): The range to be used as the data source for the chart. The range may refer to any sheet of the Calc document.
                columnheader (bool, optional):  When True, the topmost row of the range is used as labels for the category axis or the legend. Defaults to False.
                rowheader (bool, optional): When True, the leftmost column of the range is used as labels for the category axis or the legend. Defaults to False.

            Returns:
                SFDocuments.SF_Chart: A new chart service instance.

            See Also:
                `SF_Calc Help CreateChart <https://tinyurl.com/y7jwr7b7#CreateChart>`_
            """
        def DAvg(self, range: str) -> float:
            """
            Get the average of the numeric values stored in the given range

            Args:
                range (str): The range as a string where to get the values from.

            Returns:
                float: The average of the numeric values.

            See Also:
                `SF_Calc Help DAvg <https://tinyurl.com/y7jwr7b7#DAvg>`_
            """
        def DCount(self, range: str) -> float:
            """
            Get the number of numeric values stored in the given range.

            Args:
                range (str): The range as a string where to get the values from.

            Returns:
                float: The number of numeric values.

            See Also:
                `SF_Calc Help DCount <https://tinyurl.com/y7jwr7b7#DAvg>`_
            """
        def DMax(self, range: str) -> float:
            """
            Get the greatest of the numeric values stored in the given range.

            Args:
                range (str): The range as a string where to get the values from.

            Returns:
                float: The greatest of the numeric values.

            See Also:
                `SF_Calc Help DMax <https://tinyurl.com/y7jwr7b7#DAvg>`_
            """
        def DMin(self, range: str) -> float:
            """
            Get the smallest of the numeric values stored in the given range.

            Args:
                range (str): The range as a string where to get the values from.

            Returns:
                float: The smallest of the numeric values.

            See Also:
                `SF_Calc Help DMin <https://tinyurl.com/y7jwr7b7#DAvg>`_
            """
        def DSum(self, range: str) -> float:
            """
            Get sum of the numeric values stored in the given range.

            Args:
                range (str): The range as a string where to get the values from.

            Returns:
                float: The sum of the numeric values.

            See Also:
                `SF_Calc Help DSum <https://tinyurl.com/y7jwr7b7#DAvg>`_
            """
        @overload
        def Forms(self, sheetname: str) -> Tuple[str]:
            """
            Return a tuple of str with the names of all the forms contained in a given sheet.

            Args:
                sheetname (str): The name of the sheet, as a string, from which the form will be retrieved.

            Returns:
                Tuple[str]: names
                
            See Also:
                `SF_Calc Help Forms <https://tinyurl.com/y7jwr7b7#Forms>`_
            """
        @overload
        def Forms(self, sheetname: str, form: int) -> SFDocuments.SF_Form:
            """
            Retruns a SFDocuments.Form service instance representing the form specified as argument.

            Args:
                sheetname (str): The name of the sheet, as a string, from which the form will be retrieved.
                form (int): The index corresponding to a form stored in the specified sheet.

            Returns:
                SFDocuments.SF_Form: Form instance.
                
            See Also:
                `SF_Calc Help Forms <https://tinyurl.com/y7jwr7b7#Forms>`_
            """
        @overload
        def Forms(self, sheetname: str, form: str) -> SFDocuments.SF_Form:
            """
            Retruns a SFDocuments.Form service instance representing the form specified as argument.

            Args:
                sheetname (str): The name of the sheet, as a string, from which the form will be retrieved.
                form (int): The name corresponding to a form stored in the specified sheet.

            Returns:
                SFDocuments.SF_Form: Form instance.
                
            See Also:
                `SF_Calc Help Forms <https://tinyurl.com/y7jwr7b7#Forms>`_
            """
        def GetColumnName(self, columnnumber: int) -> str:
            """
            Converts a column number ranging between 1 and 1024 into its corresponding
            letter (column 'A', 'B', ..., 'AMJ'). If the given column number is outside
            the allowed range, a zero-length string is returned.

            Args:
                columnnumber (int): The column number as an integer value in the interval 1 ... 1024.

            Returns:
                str: Converts a column number converted to string.

            Note:
                The maximum number of columns allowed on a Calc sheet is 1024.
 
            See Also:
                `SF_Calc Help GetColumnName <https://tinyurl.com/y7jwr7b7#GetColumnName>`_
            """
        def GetFormula(
            self, range: str
        ) -> str | Tuple[str, str] | Tuple[Tuple[str, ...], ...]:
            """
            Get the formula(s) stored in the given range of cells as a single string, a 1D or a 2D tuple of strings.

            Args:
                range (str): The range where to get the formulas from, as a string.

            Returns:
                str | Tuple[str, str] | Tuple[Tuple[str, ...], ...]: formula
 
            See Also:
                `SF_Calc Help GetFormula <https://tinyurl.com/y7jwr7b7#GetFormula>`_
            """
        def GetValue(self, range: str) -> Any:
            """
            Get the value(s) stored in the given range of cells as a single value, a 1D array or a 2D array. All values are either doubles or strings.

            Args:
                range (str): The range where to get the values from, as a string.

            Returns:
                Any:  Get the value(s) stored in the given range of cells.
 
            See Also:
                `SF_Calc Help GetValue <https://tinyurl.com/y7jwr7b7#GetValue>`_
            """
        def ImportFromCSVFile(
            self, filename: str, destinationcell: str, filteroptions: str = ...
        ) -> str:
            """
            Imports the contents of a CSV-formatted text file and places it on a given destination cell.
            
            The destination area is cleared of all contents and formats before inserting the contents of the CSV file. The size of the modified area is fully determined by the contents of the input file.
            
            The method returns a string representing the modified range of cells.

            Args:
                filename (str): Identifies the file to open. It must follow the SF_FileSystem.FileNaming notation.
                destinationcell (str): The destination cell to insert the imported data, as a string. If instead a range is given, only its top-left cell is considered.
                filteroptions (str, optional): The arguments for the CSV input filter.

            Returns:
                str: A string representing the modified range of cells.
                The modified area depends only on the content of the source file.
            
            Note:
                Default ``filteroptions`` makes the folowing assumptions:
                    - The input file encoding is UTF8.
                    - The field separator is a comma, a semi-colon or a Tab character.
                    - The string delimiter is the double quote (").
                    - All lines are included.
                    - Quoted strings are formatted as text.
                    - Special numbers are detected.
                    - All columns are presumed to be texts, except if recognized as valid numbers.
                    - The language is English/US, which implies that the decimal separator is "." and the thousands separator is ",".
 
            See Also:
                `SF_Calc Help ImportFromCSVFile <https://tinyurl.com/y7jwr7b7#ImportFromCSVFile>`_
            """
        def ImportFromDatabase(
            self,
            filename: str = ...,
            registrationname: str = ...,
            destinationcell: str = ...,
            sqlcommand: str = ...,
            directsql: bool = ...,
        ) -> None:
            """
            Imports the contents of a database table, query or resultset, i.e. the result of a
            SELECT SQL command, inserting it on a destination cell.
            
            The destination area is cleared of all contents and formats before inserting the
            imported contents. The size of the modified area is fully determined by the contents in the table or query.

            Args:
                filename (str, optional): Identifies the file to open. It must follow the SF_FileSystem.FileNaming notation.
                registrationname (str, optional): The name to use to find the database in the databases register.
                    This argument is ignored if a filename is provided.
                destinationcell (str, optional): The destination of the imported data, as a string.
                    If a range is given, only its top-left cell is considered.
                sqlcommand (str, optional): A table or query name (without surrounding quotes or square brackets)
                    or a SELECT SQL statement in which table and field names may be surrounded by square brackets
                    or quotes to improve its readability.
                directsql (bool, optional): When True, the SQL command is sent to the database engine without
                    pre-analysis. Default is False. The argument is ignored for tables.
                    For queries, the applied option is the one set when the query was defined. Defaults to False.
 
            See Also:
                `SF_Calc Help ImportFromDatabase <https://tinyurl.com/y7jwr7b7#ImportFromDatabase>`_
            """
        @overload
        def InsertSheet(self, sheetname: str) -> bool:
            """
            Inserts a new empty sheet at the end of the list of sheets.

            Args:
                sheetname (str): The name of the new sheet.

            Returns:
                bool: True if the sheet could be inserted successfully.

            See Also:
                `SF_Calc Help InsertSheet <https://tinyurl.com/y7jwr7b7#InsertSheet>`_

            """
        @overload
        def InsertSheet(self, sheetname: str, beforesheet: int) -> bool:
            """
            Inserts a new empty sheet before an existing sheet or at the end of the list of sheets.

            Args:
                sheetname (str): The name of the new sheet.
                beforesheet (int): The index (numeric, starting from 1) of the sheet
                    before which to insert the new sheet.

            Returns:
                bool: True if the sheet could be inserted successfully.
            See Also:
                `SF_Calc Help InsertSheet <https://tinyurl.com/y7jwr7b7#InsertSheet>`_
            """
        @overload
        def InsertSheet(self, sheetname: str, beforesheet: str) -> bool:
            """
            Inserts a new empty sheet before an existing sheet or at the end of the list of sheets.

            Args:
                sheetname (str): The name of the new sheet.
                beforesheet (str): The name of the sheet before which to insert the new sheet.
                    This argument is optional and the default behavior is to insert the sheet at the last position.

            Returns:
                bool: True if the sheet could be inserted successfully.
            See Also:
                `SF_Calc Help InsertSheet <https://tinyurl.com/y7jwr7b7#InsertSheet>`_
            """
        def MoveRange(self, source: str, destination: str) -> str:
            """
            Moves a specified source range to a destination range of cells.
            The method returns a string representing the modified range of cells.
            The dimension of the modified area is fully determined by the size of the source area.

            Args:
                source (str): The source range of cells, as a string.
                destination (str): The destination cell, as a string. If a range is given,
                    its top-left cell is considered as the destination.

            Returns:
                str: A string representing the modified range of cells.
                The modified area depends only on the size of the source area.

            See Also:
                `SF_Calc Help MoveRange <https://tinyurl.com/y7jwr7b7#MoveRange>`_
            """
        @overload
        def MoveSheet(self, sheetname: str) -> bool:
            """
            Moves an existing sheet and places it at the end of the list of sheets.

            Args:
                sheetname (str): The name of the sheet to move. The sheet must exist or an exception is raised.
  
            Returns:
                bool: True if the sheet could be moved successfully.

            See Also:
                `SF_Calc Help MoveSheet <https://tinyurl.com/y7jwr7b7#MoveSheet>`_
            """
        @overload
        def MoveSheet(self, sheetname: str, beforesheet: int) -> bool:
            """
            Moves an existing sheet and places it before a specified sheet.

            Args:
                sheetname (str): The name of the sheet to move. The sheet must exist or an exception is raised.
                beforesheet (int): The index (numeric, starting from 1) of the sheet
                    before which the original sheet will be placed.

            Returns:
                bool: True if the sheet could be moved successfully.

            See Also:
                `SF_Calc Help MoveSheet <https://tinyurl.com/y7jwr7b7#MoveSheet>`_
            """
        @overload
        def MoveSheet(self, sheetname: str, beforesheet: str) -> bool:
            """
            Moves an existing sheet and places it before a specified sheet.

            Args:
                sheetname (str): The name of the sheet to move. The sheet must exist or an exception is raised.
                beforesheet (str): The name of the sheet before which the original sheet will be placed.

            Returns:
                bool: True if the sheet could be moved successfully.

            See Also:
                `SF_Calc Help MoveSheet <https://tinyurl.com/y7jwr7b7#MoveSheet>`_
            """
        def Offset(
            self,
            range: str,
            rows: int = ...,
            columns: int = ...,
            height: int = ...,
            width: int = ...,
        ) -> str:
            """
            Returns a new range (as a string) offset by a certain number of rows and columns from a given range.
            
            This method has the same behavior as the homonymous Calc's Offset function.

            Args:
                range (str): The range, as a string, that the method will use as reference to perform the offset operation.
                rows (int, optional): The number of rows by which the initial range is offset upwards (negative value) or downwards (positive value). Use 0 (default) to stay in the same row.
                columns (int, optional): The number of columns by which the initial range is offset to the left (negative value) or to the right (positive value). Use 0 (default) to stay in the same column.
                height (int, optional): The vertical height for an area that starts at the new range position. Omit this argument when no vertical resizing is needed.
                width (int, optional): The horizontal width for an area that starts at the new range position. Omit this argument when no horizontal resizing is needed.

            Returns:
                str: A new range.

            Note:
                Arguments ``rows`` and ``columns`` must not lead to zero or negative start row or column.
                
                Arguments ``height`` and ``width`` must not lead to zero or negative count of rows or columns.

            See Also:
                `SF_Calc Help Offset <https://tinyurl.com/y7jwr7b7#Offset>`_
            """
        def OpenRangeSelector(
            self,
            title: str = ...,
            selection: str = ...,
            singlecell: bool = ...,
            closeafterselect: bool = ...,
        ) -> str:
            """
            Opens a non-modal dialog that can be used to select a range in the document and returns a string containing the selected range.

            Args:
                title (str, optional): The title of the dialog, as a string.
                selection (bool, optional): An optional range that is initially
                    selected when the dialog is displayed.
                singlecell (str, optional): When True (default) only single-cell
                    selection is allowed. When False range selection is allowed.
                closeafterselect (bool, optional): When ``True`` (default) the dialog
                    is closed immediately after the selection is made.
                    When False the user can change the selection as many times as needed
                    and then manually close the dialog.

            Returns:
                str: The selected range as a string, or the empty string when the user
                cancelled the request (close window button)

            Note:
                This method opens the same dialog that is used by LibreOffice when the Shrink
                button is pressed. For example, the Tools - Goal Seek dialog has a Shrink button
                to the right of the Formula cell field.

            See Also:
                `SF_Calc Help OpenRangeSelector <https://tinyurl.com/y7jwr7b7#OpenRangeSelector>`_
            """
        @overload
        def Printf(self, inputstr: str, range: str) -> str: ...
        @overload
        def Printf(self, inputstr: str, range: str, tokencharacter: str) -> str:
            """
            Returns the input string after substituting its token characters by their values in a given range.

            Args:
                inputstr (str): The string containing the tokens that will be replaced by the corresponding values in range.
                range (str): A RangeName from which values will be extracted. If it contains a sheet name, the sheet must exist.
                tokencharacter (str): Character used to identify tokens. By default "%" is the token character.

            Note:
                Acceptable ``tokencharacter``:
                
                    - ``%S`` - The sheet name containing the range, including single quotes when necessary.
                    - ``%R1`` - The row number of the top left cell of the range.
                    - ``%C1`` - The column letter of the top left cell of the range.    
                    - ``%R2`` - The row number of the bottom right cell of the range.
                    - ``%C2`` - The column letter of the bottom right cell of the range.

            Returns:
                str: The input string after substitution of the contained tokens.

            See Also:
                `SF_Calc Help Printf <https://tinyurl.com/y7jwr7b7#Printf>`_
            """
        def PrintOut(
            self, sheetname: str = ..., pages: int = ..., copies: int = ...
        ) -> bool:
            """
            This method sends the contents of the given sheet to the default printer or to the
            printer defined by the SetPrinter method of the Document service.

            Args:
                sheetname (str, optional):The sheet to print, default is the active sheet.
                pages (int, optional): The pages to print as a string, like in the user interface.
                    Example: "1-4;10;15-18". Default is all pages.
                copies (int, optional): The number of copies. Default is 1.

            Returns:
                bool: True if the sheet was successfully printed.

            See Also:
                `SF_Calc Help PrintOut <https://tinyurl.com/y7jwr7b7#PrintOut>`_
            """
        def RemoveSheet(self, sheetname: str) -> bool:
            """
            Removes an existing sheet from the document.

            Args:
                sheetname (str): The name of the sheet to remove.

            Returns:
                bool: True if the sheet could be removed successfully.

            See Also:
                `SF_Calc Help RemoveSheet <https://tinyurl.com/y7jwr7b7#RemoveSheet>`_
            """
        def RenameSheet(self, sheetname: str, newname: str) -> bool:
            """
            Renames the given sheet/

            Args:
                sheetname (str): The name of the sheet to rename.
                newname (str): the new name of the sheet. It must not exist yet.

            Returns:
                bool: True if successful.

            See Also:
                `SF_Calc Help RenameSheet <https://tinyurl.com/y7jwr7b7#RenameSheet>`_
            """
        def SetArray(self, targetcell: str, value: Any) -> str:
            """
            Stores the given value starting from a specified target cell. The updated area expands
            itself from the target cell or from the top-left corner of the given range to accommodate
            the size of the input value argument. Vectors are always expanded vertically.

            Args:
                targetcell (str): The cell or a range as a string from where to start to store the given value.
                value (Any): A scalar, a vector or an array (in Python, one or two-dimensional lists and tuples)
                    with the new values to be stored from the target cell or from the top-left corner of the
                    range if targetcell is a range. The new values must be strings, numeric values or dates.
                    Other types will cause the corresponding cells to be emptied.

            Returns:
                str: A string representing the modified area as a range of cells.

            See Also:
                `SF_Calc Help SetArray <https://tinyurl.com/y7jwr7b7#SetArray>`_
            """
        def SetCellStyle(self, targetrange: str, style: str) -> str:
            """
            Applies the specified cell style to the given target range.
            The full range is updated and the remainder of the sheet is left untouched.
            If the cell style does not exist, an error is raised.

            Args:
                targetrange (str): The range to which the style will be applied, as a string.
                style (str): The name of the cell style to apply.

            Returns:
                str: A string representing the modified area as a range of cells.

            See Also:
                `SF_Calc Help SetCellStyle <https://tinyurl.com/y7jwr7b7#SetCellStyle>`_
            """
        def SetFormula(self, targetrange: str, formula: str) -> str:
            """
            Inserts the given (array of) formula(s) in the specified range.
            The size of the modified area is equal to the size of the range.

            Args:
                targetrange (str): The range to insert the formulas, as a string.
                formula (str): A string, a vector or an array of strings with the new
                    formulas for each cell in the target range.

            Returns:
                str: A string representing the modified area as a range of cells.

            Note:
                The full range is updated and the remainder of the sheet is left unchanged.

                If the given formula is a string, the unique formula is pasted along the
                whole range with adjustment of the relative references.

                If the size of formula is smaller than the size of ``targetrange``, then the
                remaining cells are emptied.

                If the size of formula is larger than the size of ``targetrange``, then the
                formulas are only partially copied until it fills the size of ``targetrange``.

                Vectors are always expanded vertically, except if ``targetrange`` has a height
                of exactly 1 row.

            See Also:
                `SF_Calc Help SetFormula <https://tinyurl.com/y7jwr7b7#SetFormula>`_
            """
        def SetValue(self, targetrange: str, value: str) -> str:
            """
            Stores the given value in the specified range. The size of the modified area
            is equal to the size of the target range.

            Args:
                targetrange (str): The range where to store the given value, as a string.
                value (str): A scalar, a vector or an array with the new values for each
                    cell of the range. The new values must be strings, numeric values or
                    dates. Other types will cause the corresponding cells to be emptied.

            Returns:
                str: A string representing the modified area as a range of cells.

            See Also:
                `SF_Calc Help SetValue <https://tinyurl.com/y7jwr7b7#SetValue>`_
            """
        def ShiftDown(self, range: str, wholerow: bool = ..., rows: int = ...) -> str:
            """
            Moves a given range of cells downwards by inserting empty rows.
            The current selection is not affected.
            
            Depending on the value of the wholerows argument the inserted rows can either
            span the width of the specified range or span all columns in the row.

            Args:
                range (str): The range above which rows will be inserted, as a string.
                wholerow (bool, optional): If set to False (default), then the width
                    of the inserted rows will be the same as the width of the specified range.
                    Otherwise, the inserted row will span all columns in the sheet.
                rows (int, optional): The number of rows to be inserted. The default value is
                    the height of the original range. The number of rows must be a positive number.

            Returns:
                str: a string representing the new location of the initial range.

            See Also:
                `SF_Calc Help ShiftDown <https://tinyurl.com/y7jwr7b7#ShiftDown>`_
            """
        def ShiftLeft(
            self, range: str, wholecolumn: bool = ..., columns: int = ...
        ) -> str:
            """
            Deletes the leftmost columns of a given range and moves to the left all cells to the right
            of the affected range. The current selection is not affected.
            
            Depending on the value of the wholecolumns argument the deleted columns can either span the
            height of the specified range or span all rows in the column.

            Args:
                range (str): The range from which cells will be deleted, as a string.
                wholecolumn (bool, optional): If set to False (default), then the height of the deleted
                    columns will be the same as the height of the specified range. Otherwise, the deleted
                        columns will span all rows in the sheet.
                columns (int, optional): The number of columns to be deleted from the specified range.
                    The default value is the width of the original range, which is also the maximum
                    value of this argument.

            Returns:
                str: a string representing the location of the remaining portion of the initial range.
                    If all cells in the original range have been deleted, then an empty string is returned.

            See Also:
                `SF_Calc Help ShiftLeft <https://tinyurl.com/y7jwr7b7#ShiftLeft>`_
            """
        def ShiftRight(
            self, range: str, wholecolumn: int = ..., columns: int = 000
        ) -> str:
            """
            Moves a given range of cells to the right by inserting empty columns.
            The current selection is not affected.
            
            Depending on the value of the wholecolumns argument the inserted columns can
            either span the height of the specified range or span all rows in the column.

            Args:
                range (str): The range which will have empty columns inserted to its left, as a string.
                wholecolumn (int, optional): If set to False (default), then the height of the inserted
                    columns will be the same as the height of the specified range. Otherwise, the inserted
                    columns will span all rows in the sheet.
                columns (int, optional): The number of columns to be inserted.
                    The default value is the width of the original range.

            Returns:
                str: a string representing the new location of the initial range.

            Note:
                If the shifted range exceeds the sheet edges, then nothing happens.

            See Also:
                `SF_Calc Help ShiftRight <https://tinyurl.com/y7jwr7b7#ShiftRight>`_
            """
        def ShiftUp(self, range: str, wholerow: int = ..., rows: str = ...) -> str:
            """
            Deletes the topmost rows of a given range and moves upwards all cells below the affected range.
            The current selection is not affected.
            
            Depending on the value of the wholerows argument the deleted rows can either span the width
            of the specified range or span all columns in the row.

            Args:
                range (str): The range from which cells will be deleted, as a string.
                wholerow (int, optional): If set to False (default), then the width of the deleted
                    rows will be the same as the width of the specified range. Otherwise, the deleted
                    row will span all columns in the sheet.
                rows (str, optional): The number of rows to be deleted from the specified range.
                    The default value is the height of the original range, which is also the maximum
                    value of this argument.

            Returns:
                str: A string representing the location of the remaining portion of the initial range.
                    If all cells in the original range have been deleted, then an empty string is returned.

            See Also:
                `SF_Calc Help ShiftUp <https://tinyurl.com/y7jwr7b7#ShiftUp>`_
            """
        def SortRange(
            self,
            range: str,
            sortkeys: Any,
            sortorder: Any = ...,
            destinationcell: str = ...,
            containsheader: bool = ...,
            casesensitive: bool = ...,
            sortcolumns: bool = ...,
        ) -> str:
            """
            Sorts the given range based on up to 3 columns/rows. The sorting order may vary by column/row.
            It returns a string representing the modified range of cells. The size of the modified area
            is fully determined by the size of the source area.

            Args:
                range (str): The range to be sorted, as a string.
                sortkeys (Any): A scalar (if 1 column/row) or an array of column/row numbers starting from 1.
                    The maximum number of keys is 3.
                sortorder (Any, optional): A scalar or an array of strings containing the values
                    "ASC" (ascending), "DESC" (descending) or "" (which defaults to ascending).
                    Each item is paired with the corresponding item in sortkeys.
                    If the sortorder array is shorter than sortkeys, the remaining keys are sorted in ascending order.
                destinationcell (str, optional): The destination cell of the sorted range of cells,
                    as a string. If a range is given, only its top-left cell is considered.
                    By default the source Range is overwritten.
                containsheader (bool, optional): When True, the first row/column is not sorted.
                casesensitive (bool, optional): Only for string comparisons. Default is False.
                sortcolumns (bool, optional):  When True, the columns are sorted from left to right.
                    Default is False : rows are sorted from top to bottom.

            Returns:
                str: A string representing the modified range of cells.

            See Also:
                `SF_Calc Help SortRange <https://tinyurl.com/y7jwr7b7#SortRange>`_
            """
        # endregion methods

        # region Properties
        @property
        def CurrentSelection(self) -> str | Tuple[str, ...]:
            """
            Gets/Sets the single selected range as a string or the list of selected ranges.
            """
        @property
        def Sheets(self) -> Tuple[str, ...]:
            """
            Gets a tuple with the names of all existing sheets.
            """
        # endregion Properties
    # endregion SF_Calc CLASS
    
    # region SF_CalcReference CLASS
    class SF_CalcReference(SFServices):
        """
        The SF_CalcReference class has as unique role to hold sheet and range references.
        They are implemented in Basic as Type ... End Type data structures
        """
    # endregion SF_CalcReference CLASS
    
    # region SF_Chart CLASS
    class SF_Chart(SFServices):
        """
        The SF_Chart module is focused on the description of chart documents
        stored in Calc sheets.
        With this service, many chart types and chart characteristics available
        in the user interface can be read or modified.
        
        See Also:
           `SF_Chart Help <https://tinyurl.com/ydcexzky>`_
        """
        # region Methods
        def Resize(
            self, xpos: int = ..., ypos: int = ..., width: int = ..., height: int = ...
        ) -> bool:
            """
            Changes the position of the chart in the current sheet and modifies its width and height.

            Args:
                xpos (int, optional): Specify the new ``X`` position of the chart.
                    If argument is omitted or if negative values are provided,
                    the corresponding position is left unchanged.
                ypos (int, optional): Specify the new ``y`` position of the chart.
                    If argument is omitted or if negative values are provided,
                    the corresponding position is left unchanged.
                width (int, optional): Specify the new width of the chart.
                    If this argument is omitted or if a negative value is provided,
                    the chart width is left unchanged.
                height (int, optional): Specify the new height of the chart.
                    If this argument is omitted or if a negative value is provided,
                    the chart height is left unchanged.

            Returns:
                bool: ``True`` if resizing was successful.

            Note:
                All arguments are provided as integer values that correspond to ``1/100`` of a millimeter.

            See Also:
                `SF_Chart Help Resize <https://tinyurl.com/ydcexzky#Resize>`_
            """
        def ExportToFile(
            self, filename: str, imagetype: str = ..., overwrite: bool = ...
        ) -> bool:
            """
            Saves the chart as an image file in a specified location.

            Args:
                filename (str): Identifies the path and file name where the image will be saved.
                    It must follow the notation defined in SF_FileSystem.FileNaming.
                imagetype (str, optional): The name of the image type to be created.
                overwrite (bool, optional): Specifies if the destination file can be overwritten. Defaults to False.

            Note:
                Arg ``imagetype`` accepted values:
                    - gif
                    - jpeg
                    - png (default)
                    - svg
                    - tiff

            Returns:
                bool: True if the image file could be successfully created.

            See Also:
                `SF_Chart Help ExportToFile <https://tinyurl.com/ydcexzky#ExportToFile>`_
            """
        # endregion Methods

        # region Properties
        @property
        def ChartType(self) -> str:
            """
            Gets/Sets the chart type as a string that can assume one of the following values:
            
            Pie, Bar, Donut, Column, Area, Line, XY, Bubble, Net.
            """
        @property
        def Deep(self) -> bool:
            """
            Gets/Sets if the chart is three-dimensional.
            
            When True indicates that the chart is three-dimensional and each series is arranged in the z-direction.
            
            When False series are arranged considering only two dimensions.
            """
        
        @property
        def Dim3D(self) -> bool | str:
            """
            Gets/Sets if the chart is displayed with 3D elements.
            
            When setting as string the value must be either "Bar", "Cylinder", "Cone" or "Pyramid".
            
            When True value is specified, then the chart is displayed using 3D bars.
            """
        @property
        def Exploded(self) -> float:
            """
            Gets/Sets how much pie segments are offset from the chart center as a percentage of the radius.
            
            Applicable to pie and donut charts only.
            """
        @property
        def Filled(self) -> bool:
            """
            Gets/Sets if a filled net chart
            
            When True, specifies a filled net chart.
            
            Applicable to net charts only.
            """
        @property
        def Legend(self) -> bool:
            """
            Gets/Sets whether or not the chart has a legend.
            """
        @property
        def Percent(self) -> bool:
            """
            Gets/Sets percent
            
            When True, chart series are stacked and each category sums up to 100%.
            
            Applicable to Area, Bar, Bubble, Column and Net charts.
            """ 
        @property
        def Stacked(self) -> bool:
            """
            Gets/Sets whether or not the chart series are stacked.
            
            When True, chart series are stacked and each category sums up to 100%.
            
            Applicable to Area, Bar, Bubble, Column and Net charts.
            """ 
        @property
        def Title(self) -> str:
            """Gets/Sets the main title of the chart."""
        @property
        def XTitle(self) -> str:
            """"Gets/Sets the title of the X axis."""
        @property
        def YTitle(self) -> str:
            """"Gets/Sets title of the Y axis."""
        @property
        def XChartObj(self) -> object:
            """
            Gets the object representing the chart, which is an instance of the ScChartObj class.
            
            See Also:
                `ScChartObj <https://docs.libreoffice.org/sc/html/classScChartObj.html>`_
            """
        @property
        def XDiagram(self) -> XDiagram:
            """
            Gets the ``com.sun.star.chart.XDiagram`` object representing the diagram of the chart.
            """
        @property
        def XShape(self) -> XShape:
            """
            Gets the com.sun.star.drawing.XShape object representing the shape of the chart.
            """
        @property
        def XTableChart(self) -> XTableChart:
            """
            Gets the com.sun.star.table.XTableChart object representing the data being displayed in the chart.
            """
        # endregion Properties
    # endregion SF_Chart CLASS
    
    # region SF_Form CLASS
    class SF_Form(SFServices):
        """
        Management of forms defined in LibreOffice documents. Supported types are Base, Calc and Writer documents.
        It includes the management of subforms
        Each instance of the current class represents a single form or a single subform
        A form may optionally be (understand "is often") linked to a data source manageable with
        the SFDatabases.Database service. The current service offers a rapid access to that service.

        See Also:
            `SF_Form Help <https://tinyurl.com/y72zdzjy>`_
        """
        # region Methods
        def Activate(self) -> bool:
            """
            Sets the focus on the current Form instance. Returns True if focusing was successful.
            
            The behavior of the Activate method depends on the type of document where the form is located.
                * In Writer documents: Sets the focus on that document.
                * In Calc documents: Sets the focus on the sheet to which the form belongs.
                * In Base documents: Sets the focus on the FormDocument the Form refers to.

            Returns:
                bool: True if focusing was successful.

            See Also:
                `SF_Form Help Activate <https://tinyurl.com/y72zdzjy#Activate>`_
            """
        def CloseFormDocument(self) -> bool:
            """
            Closes the ``form`` document containing the actual Form instance. The ``Form`` instance is disposed.

            Returns:
                bool: True if closure is successful.

            Note:
                This method only closes form documents located in Base documents.
                If the form is stored in a Writer or Calc document, calling ``CloseFormDocument`` will have no effect.

            See Also:
                `SF_Form Help CloseFormDocument <https://tinyurl.com/y72zdzjy#CloseFormDocument>`_
            """
        @overload
        def Controls(self) -> Tuple[str, ...]:
            """
            Gets  the list of the controls contained in the form. Beware that the returned list does not contain any subform controls.

            Returns:
                Tuple[str, ...]: The list of the controls contained in the form

            See Also:
                `SF_Form Help Controls <https://tinyurl.com/y72zdzjy#Controls>`_
            """
        @overload
        def Controls(self, controlname: str) -> "SFDocuments.SF_FormControl":
            """
            Gets a FormControl class instance referring to the specified control.

            Args:
                controlname (str): A valid control name as a case-sensitive string.

            Returns:
                SFDocuments.SF_FormControl: A FormControl class instance referring to the specified control.

            See Also:
                `SF_Form Help Controls <https://tinyurl.com/y72zdzjy#Controls>`_
            """
        def GetDatabase(
            self, user: str = ..., password: str = ...
        ) -> SFDatabases.SF_Database:
            """
            Gets a SFDatabases.Database instance giving access to the execution of SQL commands on the database the current form is connected to and/or that is stored in the current Base document.
            
            Each form has its own database connection, except in Base documents where they all share the same connection.

            Args:
                user (str, optional): The login optional parameter.
                password (str, optional): user password.

            Returns:
                SFDatabases.SF_Database: a SFDatabases.Database instance giving access to the execution of
                    SQL commands on the database the current form is connected to and/or that is stored
                    in the current Base document.

            See Also:
                `SF_Form Help GetDatabase <https://tinyurl.com/y72zdzjy#GetDatabase>`_
            """
        def MoveFirst(self) -> bool:
            """
            The form cursor is positioned on the first record.

            Returns:
                bool: True if successful.

            See Also:
                `SF_Form Help MoveFirst <https://tinyurl.com/y72zdzjy#MoveFirst>`_
            """
        def MoveLast(self) -> bool:
            """
            The form cursor is positioned on the last record.

            Returns:
                bool: True if successful.

            See Also:
                `SF_Form Help MoveLast <https://tinyurl.com/y72zdzjy#MoveLast>`_
            """
        def MoveNew(self) -> bool:
            """
            The form cursor is positioned on the new record area.

            Returns:
                bool: True if successful.

            See Also:
                `SF_Form Help MoveNew <https://tinyurl.com/y72zdzjy#MoveNew>`_
            """
        def MoveNext(self, offset: int = ...) -> bool:
            """
            The form cursor is positioned on the next record.

            Args:
                offset (int, optional): The number of records to go forward. Defaults to 1.

            Returns:
                bool: True if successful.

            See Also:
                `SF_Form Help MoveNext <https://tinyurl.com/y72zdzjy#MoveNext>`_
            """
        def MovePrevious(self, offset: int = ...) -> bool:
            """
            The form cursor is positioned on the previous record.

            Args:
                offset (int, optional): The number of records to go backwards. Defaults to 1.

            Returns:
                bool: True if successful.

            See Also:
                `SF_Form Help MovePrevious <https://tinyurl.com/y72zdzjy#MovePrevious>`_
            """
        def Requery(self) -> bool:
            """
            Reloads the current data from the database and refreshes the form.
            The cursor is positioned on the first record.

            Returns:
                bool: True if successful.

            See Also:
                `SF_Form Help Requery <https://tinyurl.com/y72zdzjy#Requery>`_
            """
        @overload
        def Subforms(self) -> Tuple[str, ...]:
            """
            Gets the list of subforms contained in the current form or subform instance.

            Returns:
                Tuple[str, ...]: List of subforms.

            See Also:
                `SF_Form Help Subforms <https://tinyurl.com/y72zdzjy#Subforms>`_
            """
        @overload
        def Subforms(self, subform: str) -> SFDocuments.SF_Form:
            """
            Gets SFDocuments.Form instance based on the specified form/subform name.

            Args:
                subform (int): A subform stored in the current Form class instance given by its name.

            Returns:
                SFDocuments.SF_Form: SFDocuments.Form instance based on the specified form/subform name.

            See Also:
                `SF_Form Help Subforms <https://tinyurl.com/y72zdzjy#Subforms>`_
            """
        @overload
        def Subforms(self, subform: int) -> SFDocuments.SF_Form:
            """
            Gets SFDocuments.Form instance based on the specified form/subform index.

            Args:
                subform (int): A subform stored in the current Form class instance given by its index.

            Returns:
                SFDocuments.SF_Form: SFDocuments.Form instance based on the specified form/subform index.

            See Also:
                `SF_Form Help Subforms <https://tinyurl.com/y72zdzjy#Subforms>`_
            """
        # endregion Methods
        
        # region Properties
        @property
        def AllowDeletes(self) -> bool:
            """
            Gets/Sets if the form allows to delete records.
            """
        @property
        def AllowInserts(self) -> bool:
            """
            Gets/Sets if the form allows to add records.
            """
        @property
        def AllowUpdates(self) -> bool:
            """
            Gets/Sets if the form allows to update records.
            """
        @property
        def BaseForm(self) -> str:
            """
            Gets the hierarchical name of the Base Form containing the actual form.
            """
        @property
        def Bookmark(self) -> Any:
            """
            Gets/Sets uniquely the current record of the form's underlying table, query or SQL statement.
            """
        @property
        def CurrentRecord(self) -> int:
            """
            Gets/Sets the current record in the dataset being viewed on a form.
            
            If the row number is positive, the cursor moves to the given row number
            with respect to the beginning of the result set. Row count starts at 1.
            If the given row number is negative, the cursor moves to an absolute
            row position with respect to the end of the result set.
            Row -1 refers to the last row in the result set.
            """
        @property
        def Filter(self) -> bool:
            """
            Gets/Sets filter for subset.
            
            Specifies a subset of records to be displayed as a
            SQL WHERE-clause without the WHERE keyword.
            """
        @property
        def LinkChildFields(self) -> str:
            """
            Gets how records in a child subform are linked to records in its parent form.
            """
        @property
        def LinkParentFields(self) -> str:
            """
            Gets how records in a child subform are linked to records in its parent form.
            """
        @property
        def Name(self) -> str:
            """
            Gets the name of the current form.
            """
        @property
        def OnApproveCursorMove(self) -> str:
            """
            Gets/Sets URI strings that define the script triggered by the event.
            """
        @property
        def OnApproveParameter(self) -> str:
            """
            Gets/Sets URI strings that define the script triggered by the event.
            """
        @property
        def OnApproveReset(self) -> str:
            """
            Gets/Sets URI strings that define the script triggered by the event.
            """
        @property
        def OnApproveRowChange(self) -> str:
            """
            Gets/Sets URI strings that define the script triggered by the event.
            """
        @property
        def OnApproveSubmit(self) -> str:
            """
            Gets/Sets URI strings that define the script triggered by the event.
            """
        @property
        def OnConfirmDelete(self) -> str:
            """
            Gets/Sets URI strings that define the script triggered by the event.
            """
        @property
        def OnCursorMoved(self) -> str:
            """
            Gets/Sets URI strings that define the script triggered by the event.
            """
        @property
        def OnErrorOccurred(self) -> str:
            """
            Gets/Sets URI strings that define the script triggered by the event.
            """
        @property
        def OnLoaded(self) -> str:
            """
            Gets/Sets URI strings that define the script triggered by the event.
            """
        @property
        def OnReloaded(self) -> str:
            """
            Gets/Sets URI strings that define the script triggered by the event.
            """
        @property
        def OnReloading(self) -> str:
            """
            Gets/Sets URI strings that define the script triggered by the event.
            """
        @property
        def OnResetted(self) -> str:
            """
            Gets/Sets URI strings that define the script triggered by the event.
            """
        @property
        def OnRowChanged(self) -> str:
            """
            Gets/Sets URI strings that define the script triggered by the event.
            """
        @property
        def OnUnloaded(self) -> str:
            """
            Gets/Sets URI strings that define the script triggered by the event.
            """
        @property
        def OnUnloading(self) -> str:
            """
            Gets/Sets URI strings that define the script triggered by the event.
            """
        @property
        def OrderBy(self) -> str:
            """
            Gets/Sets in which order the records should be displayed
            as a SQL ORDER BY clause without the ORDER BY keywords.
            """
        @property
        def Parent(self) -> SFDocuments.SF_Form | SFDocuments.SF_Document:
            """
            Gets the parent of the current form. It can be either a SFDocuments.Form or a SFDocuments.Document object.
            """
        @property
        def RecordSource(self) -> str:
            """
            Gets/Sets the source of the data, as a table name, a query name or a SQL statement.
            """
        @property
        def XForm(self) -> XForm:
            """
            Gets The UNO object representing interactions with the form.
            
            Refer to XForm and DataForm in the API documentation for detailed information.
            """
        # endregion Properties
    # endregion SF_Form CLASS
    
    # region SF_FormControl CLASS
    class SF_FormControl(SFServices):
        """
        Manage the controls belonging to a form or subform stored in a document.
        Each instance of the current class represents a single control within a form, a subform or a tablecontrol.
        A prerequisite is that all controls within the same form, subform or tablecontrol must have
        a unique name.

        See Also:
            `SF_FormControl Help <https://tinyurl.com/y8d9qlcl>`_
        """
        # region Methods
        @overload
        def Controls(self) -> Tuple[str, ...]:
            """
            Gets tuple containing the names of all controls

            Returns:
                Tuple[str, ...]: tuple containing the names of all controls.

            See Also:
                `SF_FormControl Help Controls <https://tinyurl.com/y8d9qlcl#Controls>`_
            """
        @overload
        def Controls(self, controlname: str) -> SFDocuments.SF_FormControl:
            """
            Gets a FormControl class instance corresponding to the specified control.

            Args:
                controlname (str): A valid control name as a case-sensitive string.

            Returns:
                SFDocuments.SF_FormControl: A FormControl class instance corresponding to the specified control.

            See Also:
                `SF_FormControl Help Controls <https://tinyurl.com/y8d9qlcl#Controls>`_
            """
        def SetFocus(self) -> bool:
            """
            Sets the focus on the control.

            Returns:
                bool: True if focusing was successful.

            See Also:
                `SF_FormControl Help SetFocus <https://tinyurl.com/y8d9qlcl#SetFocus>`_
            """
        # endregion Methods
        
        # region Properties
        @property
        def Action(self) -> str:
            """
            Gets/Sets the action triggered when the button is clicked.
            
            Accepted values are:
                - none
                - submitForm
                - resetForm
                - refreshForm
                - moveToFirst
                - moveToLast
                - moveToNext
                - moveToPrev
                - saveRecord
                - moveToNew
                - deleteRecord
                - undoRecord

            Applicable Controls:
                Button
            """
        @property
        def Action(self) -> str:
            """
            Gets/Sets the text displayed by the control.

            Applicable Controls:
                - Button
                - CheckBox
                - FixedText
                - GroupBox
                - RadioButton
            """
        @property
        def ControlSource(self) -> str:
            """
            Gets the rowset field mapped onto the current control.
            
            Applicable Controls:
                - CheckBox
                - ComboBox
                - CurrencyField
                - DateField
                - FormattedField
                - ImageControl
                - ListBox
                - NumericField
                - PatternField
                - RadioButton
                - TextField
                - TimeField
            """
        @property
        def ControlType(self) -> str:
            """
            Gets control type from one of the controls listed in ControlSource property.
            
            Applicable Controls:
                All
            """
        @property
        def Default(self) -> bool:
            """
            Gets/Sets if a command button is the default OK button.

            Applicable Controls:
                Button
            """
        @property
        def DefaultValue(self) -> Any:
            """
            Gets/Sets the default value used to initialize a control in a new record.
            
            Applicable Controls:
                - CheckBox
                - ComboBox
                - CurrencyField
                - DateField
                - FileControl
                - FormattedField
                - ListBox
                - NumericField
                - PatternField
                - RadioButton
                - SpinButton
                - TextField
                - TimeField
            """
        @property
        def Enabled(self) -> bool:
            """
            Gets/Sets if the control is accessible with the cursor.

            Applicable Controls:
                All (except HiddenControl)
            """
        @property
        def Format(self) -> str:
            """
            Gets/Sets the format used to display dates and times.
            
            Must be one of following strings for dates:
                - "Standard (short)"
                - "Standard (short YY)"
                - "Standard (short YYYY)"
                - "Standard (long)"
                - "DD/MM/YY"
                - "MM/DD/YY"
                - "YY/MM/DD"
                - "DD/MM/YYYY"
                - "MM/DD/YYYY"
                - "YYYY/MM/DD"
                - "YY-MM-DD"
                - "YYYY-MM-DD"

            Applicable Controls:
                - DateField
                - TimeField
                - FormattedField (read-only)
            """
        @property
        def ListCount(self) -> int:
            """
            Gets the number of rows in a ListBox or a ComboBox.

            Applicable Controls:
                - ComboBox
                - ListBox
            """
        @property
        def ListIndex(self) -> int:
            """
            Gets/Sets which item is selected in a ListBox or ComboBox.

            In case of multiple selection, the index of the first item
            is returned or only one item is set.

            Applicable Controls:
                - ComboBox
                - ListBox
            """
        @property
        def ListSource(self) -> ListSourceType:
            """
            Gets/Sets the type of data contained in a combobox or a listbox.

            It must be one of the com.sun.star.form.ListSourceType.* constants.

            Applicable Controls:
                - ComboBox
                - ListBox
            """
        @property
        def Locked(self) -> bool:
            """
            Gets/Sets if the control is read-only.

            Applicable Controls:
                - ComboBox
                - CurrencyField
                - DateField
                - FileControl
                - FileControl
                - FormattedField
                - ImageControl
                - ListBox
                - NumericField
                - PatternField
                - TextField
                - TimeField
            """
        @property
        def MultiSelect(self) -> bool:
            """
            Gets/Sets if the user can select multiple items in a listbox.

            Applicable Controls:
                ListBox
            """
        @property
        def Name(self) -> str:
            """
            Gets the name of the control.

            Applicable Controls:
                All
            """
        @property
        def OnActionPerformed(self) -> str:
            """
            Gets/Sets URI strings that define the script triggered by the event.
            """
        @property
        def OnAdjustmentValueChanged(self) -> str:
            """
            Gets/Sets URI strings that define the script triggered by the event.
            """
        @property
        def OnApproveAction(self) -> str:
            """
            Gets/Sets URI strings that define the script triggered by the event.
            """
        @property
        def OnApproveReset(self) -> str:
            """
            Gets/Sets URI strings that define the script triggered by the event.
            """
        @property
        def OnApproveUpdate(self) -> str:
            """
            Gets/Sets URI strings that define the script triggered by the event.
            """
        @property
        def OnChanged(self) -> str:
            """
            Gets/Sets URI strings that define the script triggered by the event.
            """
        @property
        def OnErrorOccurred(self) -> str:
            """
            Gets/Sets URI strings that define the script triggered by the event.
            """
        @property
        def OnFocusGained(self) -> str:
            """
            Gets/Sets URI strings that define the script triggered by the event.
            """
        @property
        def OnFocusLost(self) -> str:
            """
            Gets/Sets URI strings that define the script triggered by the event.
            """
        @property
        def OnItemStateChanged(self) -> str:
            """
            Gets/Sets URI strings that define the script triggered by the event.
            """
        @property
        def OnKeyPressed(self) -> str:
            """
            Gets/Sets URI strings that define the script triggered by the event.
            """
        @property
        def OnKeyReleased(self) -> str:
            """
            Gets/Sets URI strings that define the script triggered by the event.
            """
        @property
        def OnMouseDragged(self) -> str:
            """
            Gets/Sets URI strings that define the script triggered by the event.
            """
        @property
        def OnMouseEntered(self) -> str:
            """
            Gets/Sets URI strings that define the script triggered by the event.
            """
        @property
        def OnMouseExited(self) -> str:
            """
            Gets/Sets URI strings that define the script triggered by the event.
            """
        @property
        def OnMouseMoved(self) -> str:
            """
            Gets/Sets URI strings that define the script triggered by the event.
            """
        @property
        def OnMousePressed(self) -> str:
            """
            Gets/Sets URI strings that define the script triggered by the event.
            """
        @property
        def OnMouseReleased(self) -> str:
            """
            Gets/Sets URI strings that define the script triggered by the event.
            """
        @property
        def OnResetted(self) -> str:
            """
            Gets/Sets URI strings that define the script triggered by the event.
            """
        @property
        def OnTextChanged(self) -> str:
            """
            Gets/Sets URI strings that define the script triggered by the event.
            """
        @property
        def OnUpdated(self) -> str:
            """
            Gets/Sets URI strings that define the script triggered by the event.
            """
        @property
        def Parent(self) -> SFDocuments.SF_Form | SFDocuments.SF_FormControl:
            """
            Gets parent type.
            
            Depending on the parent type, a form, a subform or a tablecontrol,
            returns the parent ``SFDocuments.Form`` or ``SFDocuments.FormControl``
            class object instance.

            Applicable Controls:
                All
            """
        @property
        def Name(self) -> str:
            """
            Gets/Sets the file name containing a bitmap or other type of graphic to be displayed on the control.

            The filename must comply with the FileNaming attribute of the ``ScriptForge.FileSystem`` service.

            Applicable Controls:
                - Button
                - ImageButton
                - ImageControl
            """
        @property
        def Required(self) -> bool:
            """
            Gets/Sets if a control is said required when the underlying data must not contain a null value.

            Applicable Controls:
                - CheckBox
                - ComboBox
                - CurrencyField
                - DateField
                - ListBox
                - NumericField
                - PatternField
                - RadioButton
                - SpinButton
                - TextField
                - TimeField
            """
        @property
        def Text(self) -> str:
            """
            Gets the text being displayed by the control.

            Applicable Controls:
                - ComboBox
                - DateField
                - FileControl
                - FormattedField
                - PatternField
                - TextField
                - TimeField
            """
        @property
        def TipText(self) -> str:
            """
            Gets/Sets the text that appears as a tooltip when you hold the mouse pointer over the control.

            Applicable Controls:
                All (except HiddenControl)
            """
        @property
        def TripleState(self) -> bool:
            """
            Gets/Sets if the control may have the state "don't know".
            
            Applicable Controls:
                CheckBox
            """
        @property
        def Value(self) -> Any:
            """
            Gets/Sets control type.

            See Also:
                `The Value property <https://tinyurl.com/yb27tk36#hd_id81598540704978>`_
            """
        @property
        def Visible(self) -> bool:
            """
            Gets/Sets if the control is hidden or visible.
            
            Applicable Controls:
                All (except HiddenControl)
            """
        @property
        def XControlModel(self) -> XControlModel:
            """
            Gets the UNO object representing the control model.
            
            Applicable Controls:
                All
            """
        @property
        def XControlView(self) -> XControl:
            """
            Gets the UNO object representing the control view.
            
            Applicable Controls:
                All
            """
        # endregion Properties
    # endregion SF_FormControl CLASS
    
    # region SF_Writer CLASS
    class SF_Writer(SF_Document, SFServices):
        """
        The SF_Writer module is focused on :
            - TBD

        See Also:
            `SF_Writer Help <https://tinyurl.com/y7kv226a>`_
        """
        # region Methods
        @classmethod
        def ReviewServiceArgs(cls, windowname: str = ...) -> Tuple[str]:
            """
            Transform positional and keyword arguments into positional only
            """
        @overload
        def Forms(self) -> Tuple[str, ...]:
            """
            Gets the names of all the forms contained in the document.

            Returns:
                Tuple[str, ...]: names of all forms in document.
            
            Note:
                This method is applicable only for Writer documents.
                Calc and Base documents have their own Forms method in the
                Calc and Base services, respectively.

            See Also:
                `SF_Writer Help Forms <https://tinyurl.com/y7kv226a#Forms>`_
            """
        @overload
        def Forms(self, form: str) -> SFDocuments.SF_Form:
            """
            Gets SFDocuments.Form service instance representing the form specified as argument.

            Args:
                form (str): The name or index corresponding to a form stored in the document.

            Returns:
                SFDocuments.SF_Form: SFDocuments.Form service instance
            
            Note:
                This method is applicable only for Writer documents.
                Calc and Base documents have their own Forms method in the
                Calc and Base services, respectively.

            See Also:
                `SF_Writer Help Forms <https://tinyurl.com/y7kv226a#Forms>`_
            """
        def PrintOut(
            self,
            pages: str = ...,
            copies: int = ...,
            printbackground: bool = ...,
            printblankpages: bool = ...,
            printevenpages: bool = ...,
            printoddpages: bool = ...,
            printimages: bool = ...,
        ) -> bool:
            """
            Send the contents of the document to the printer. The printer may be previously defined by default,
            by the user or by the SetPrinter method of the Document service.

            Args:
                pages (str, optional): The pages to print as a string, like in the user interface.
                    Example: "1-4;10;15-18". Default = all pages
                copies (int, optional): The number of copies. Defaults to 1.
                printbackground (bool, optional): Prints the background image when True. Defaults to True.
                printblankpages (bool, optional): When False, omits empty pages. Defaults to False.
                printevenpages (bool, optional): Prints even pages when True. Defaults to True.
                printoddpages (bool, optional): Print odd pages when True. Defaults to True.
                printimages (bool, optional): Print graphic objects when True. Defaults to True.

            Returns:
                bool: True when successful.

            See Also:
                `SF_Writer Help PrintOut <https://tinyurl.com/y7kv226a#PrintOut>`_
            """
        # endregion Methods
    # endregion SF_Writer CLASS
# endregion SFDocuments CLASS    (alias of SFDocuments Basic library)

# region SFWidgets CLASS    (alias of SFWidgets Basic library)
class SFWidgets:
    """
    The SFWidgets class manages toolbars and popup menus
    """

    # region SF_PopupMenu CLASS
    class SF_PopupMenu(SFServices):
        """
        Display a popup menu anywhere and any time.
        A popup menu is usually triggered by a mouse action (typically a right-click) on a dialog, a form
        or one of their controls. In this case the menu will be displayed below the clicked area.
        When triggered by other events, including in the normal flow of a user script, the script should
        provide the coordinates of the topleft edge of the menu versus the actual component.
        The menu is described from top to bottom. Each menu item receives a numeric and a string identifier.
        The execute() method returns the item selected by the user.

        See Also:
            `SF_PopupMenu Help <https://tinyurl.com/y7ngmoa8>`_
        """
        # region Methods
        @classmethod
        def ReviewServiceArgs(
            cls, event: Any = ..., x: int = ..., y: int = ..., submenuchar: str = ...
        ) -> Tuple[Any, int, int, str]:
            """
            Transform positional and keyword arguments into positional only
            """
        def AddCheckBox(
            self,
            menuitem: str,
            name: str = "",
            status: bool = False,
            icon: str = "",
            tooltip: str = "",
        ) -> int:
            """
            Inserts a check box in the popup menu.

            Args:
                menuitem (str): Defines the text to be displayed in the menu.
                    This argument also defines the hierarchy of the item inside the
                    menu by using the submenu character.
                name (str, optional): String value to be returned when the item is clicked.
                    By default, the last component of the menu hierarchy is used.
                status (bool, optional): Defines whether the item is selected when the menu is created.
                    Defaults to ``False``.
                icon (str, optional): Path and name of the icon to be displayed without the leading
                    path separator. The actual icon shown depends on the icon set being used.
                tooltip (str, optional): Text to be displayed as tooltip.
            Returns:
                int: integer value that identifies the inserted item.

            See Also:
                `SF_PopupMenu Help AddCheckBox <https://tinyurl.com/y7ngmoa8#AddCheckBox>`_
            """
        def AddItem(
            self, menuitem: str, name: str = ..., icon: str = ..., tooltip: str = ...
        ) -> int:
            """
            Inserts a menu entry in the popup menu

            Args:
                menuitem (str): Defines the text to be displayed in the menu.
                    This argument also defines the hierarchy of the item inside the menu
                    by using the submenu character.
                name (str, optional): String value to be returned when the item is clicked.
                    By default, the last component of the menu hierarchy is used.
                icon (str, optional): Path and name of the icon to be displayed without the leading
                    path separator. The actual icon shown depends on the icon set being used.
                tooltip (str, optional): Text to be displayed as tooltip.

            Returns:
                int: integer value that identifies the inserted item.

            See Also:
                `SF_PopupMenu Help AddItem <https://tinyurl.com/y7ngmoa8#AddItem>`_
            """
        def AddRadioButton(
            self,
            menuitem: str,
            name: str = ...,
            status: bool = ...,
            icon: str = ...,
            tooltip: str = ...,
        ) -> int:
            """
            Inserts a radio button entry in the popup menu.

            Args:
                menuitem (str): Defines the text to be displayed in the menu.
                    This argument also defines the hierarchy of the item inside the
                    menu by using the submenu character.
                name (str, optional): String value to be returned when the item is clicked.
                    By default, the last component of the menu hierarchy is used.
                status (bool, optional): Defines whether the item is selected when the menu is created.
                    Defaults to ``False``.
                icon (str, optional): Path and name of the icon to be displayed without the leading
                    path separator. The actual icon shown depends on the icon set being used.
                tooltip (str, optional): Text to be displayed as tooltip.

            Returns:
                int: integer value that identifies the inserted item.

            See Also:
                `SF_PopupMenu Help AddRadioButton <https://tinyurl.com/y7ngmoa8#AddRadioButton>`_
            """
        def Execute(self, returnid: bool = ...) -> Union[int, str]:
            """
            Displays the popup menu and waits for a user action.
            
            If the user clicks outside the popup menu ou presses the Esc key, then no item is selected.
            In such cases, the returned value depends on the ``returnid`` parameter.
            If returnid = ``True`` and no item is selected, then the value ``0`` (zero) is returned.
            Otherwise an empty string ``''`` is returned.

            Args:
                returnid (bool, optional): If True the selected item ID is returned.
                    If False the method returns the item's name. Defaults to True.

            Returns:
                int | str: The item clicked by the user.

            See Also:
                `SF_PopupMenu Help Execute <https://tinyurl.com/y7ngmoa8#Execute>`_
            """
        # endregion Methods
        
        # region Properties
        @property
        def SubmenuCharacter(self) -> str:
            """
            Gets/Sets the character or string that defines how menu items are nested. The default character is >.
            """
        @property
        def ShortcutCharacter(self) -> str:
            """
            Gets/Sets the character used to define the access key of a menu item. The default character is ~.
            """
        # endregion Properties
    # endregion SF_PopupMenu CLASS
# endregion SFWidgets CLASS    (alias of SFWidgets Basic library)

# region CreateScriptService()
def CreateScriptService(service: str, *args: Any, **kwargs: Any) -> Union[SFServices, Any]:
    """
    A service being the name of a collection of properties and methods,
    this method returns either
        - the Python object mirror of the Basic object implementing the requested service
        - the Python object implementing the service itself

    A service may be designated by its official name, stored in its class.servicename
    or by one of its synonyms stored in its class.servicesynonyms list
    If the service is not identified, the service creation is delegated to Basic, that might raise an error
    if still not identified there

    Args:
        service (str): the name of the service as a string 'library.service' - cased exactly
            or one of its synonyms
        args (any, optional): the arguments to pass to the service constructor
    
     Returns:
        SFServices | Any:  the service as a Python object
    """
# endregion CreateScriptService()

createScriptService: Union[SFServices, Any]
createscriptservice: Union[SFServices, Any]

# ######################################################################
# Lists the scripts, that shall be visible inside the Basic/Python IDE
# ######################################################################

g_exportedScripts: tuple
