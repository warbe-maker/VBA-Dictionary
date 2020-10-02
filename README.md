## Common VBA Dictionary Procedures

Methods, useful when working with Dictionaries.

### Installation
Download [_mDct.bas_](https://gitcdn.link/repo/warbe-maker/Common-VBA-Dictionary-Procedures/master/mDct.bas) and import it into your VB-Project. Alternatively you may fork the [Github repo Common-VBA-Dictionary-Procedures](https://github.com/warbe-maker/Common-VBA-Dictionary-Procedures).

### Usage
#### Method DctAdd
See blog-post [Add key/item pairs to a Dictionary "instantly ordered"](https://warbe-maker.github.io)
#### Method DctDiff
stil to be added here

### Contribution (Development, test, maintenance)
- The module **_mDct_** is hosted in the dedicated _Common Component Workbook_ **_Dct.xlsm_** which is used as the development, test, and maintenance environment.
- In the Workbook the procedure **_Test\_DctAdd\_00\_Regression_** in module **_mTest_** provides a fully automated regression test, obligatory after any kind of code modification
- The procedure **_Test\_DctAdd\_99\_Performance_** in module **_mTest_** provides a performance test. The result for various numbers of adds can bee seen in the Test-Worksheet. For the trace of the execution time the tests in use the **_mErrHndlr_** module (not required for any procedure in the **_mDct_** module)
- The **_DctAdd_** procedure uses the **_ErrMsg\*_** procedures in module _mBasic_ which may be copied to the **_mDct_** module.

The **_mDct_** module is the ideal module for other useful procedures such like Dictionary sort which will be added soon.