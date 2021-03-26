## Common VBA Dictionary Procedures

Methods, useful when working with Dictionaries.

### Installation
Download [_mDct.bas_][1] and import it into your VB-Project. Alternatively the [Github repo Common-VBA-Dictionary-Procedures][2] may be forked or cloned.

### Usage
#### DctAdd service
Adding items to a Dictionary instantly ordered.
See blog-post [Add key/item pairs to a Dictionary "instantly ordered"](3)

#### Other services
still to be added here

### Contribution (Development, test, maintenance)
Contribution of any kind in any way is welcome.
- The module **_mDct_** is hosted in the dedicated _Common Component Workbook_ **_Dct.xlsm_** which is used as the development, test, and maintenance environment.
- In the Workbook the procedure **_Test\_DctAdd\_00\_Regression_** in module **_mTest_** provides a fully automated regression test, obligatory after any kind of code modification
- The procedure **_Test\_DctAdd\_99\_Performance_** in module **_mTest_** provides a performance test. The result for various numbers of adds can bee seen in the Test-Worksheet. For the trace of the execution time the tests in use the **_mErrHndlr_** module (not required for any procedure in the **_mDct_** module)
- The **_DctAdd_** procedure uses the **_ErrMsg\*_** procedures in module _mBasic_ which may be copied to the **_mDct_** module.

The **_mDct_** module is the potential module for other useful services such like a _Sort_ (by item/key, ascending/descending, case-sensitive/case-insensitive) service which should be added in the future.

[1]: https://gitcdn.link/repo/warbe-maker/Common-VBA-Dictionary-Procedures/master/source/mDct.bas
[2]: https://github.com/warbe-maker/Common-VBA-Dictionary-Procedures
[3]: https://warbe-maker.github.io/warbe-maker.github.io/vba/dictionary/common/2020/10/02/Common-VBA-Dictionary-instantly-ordered.html
