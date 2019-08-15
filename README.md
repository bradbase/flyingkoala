
# FlyingKoala

FlyingKoala facilitates defining models (mathematical, technical and financial), scenario analysis and some system integration tasks in MS Excel while reducing the amount of computer code required to do these tasks and increasing the speed of calculation. The outcome is that people with good Excel skills can do more for themselves before requiring a code savvy offsider. FlyingKoala also facilitates communication of financial, technical and mathematical modelling as the expressions may be largely expressed in Excel formulas.

From a technical standpoint FlyingKoala is a collection of helper functions for [xlwings](https://www.xlwings.org/). These helper functions reach out to:
* [Koala2](https://github.com/vallettea/koala/blob/master/doc/presentation.md)
* [Pandas](https://pandas.pydata.org/)
* [Harvest](https://pypi.org/project/python-harvest-redux/5.0.0b0/)
* [PVLib](https://pvlib-python.readthedocs.io/en/stable/)

To a large extent the advantages for data analysis with FlyingKoala actually comes from clever use of Excel's existing functionality and the greatness of xlwings. FlyingKoala provides a number of pre-written Python UDFs which are especially useful in going beyond the usual limits of Excel.

In particular, the addition of Koala functionality significantly improves the speed of mathematical calculation which increases opportunity to process much larger data sets than Excel can usually manage and iterate over scenarios more quickly.

Wrapping a variety of things from Pandas is also a great effort in terms of time series data analysis.


# The problem space

* Auditing is difficult when everything is coded in code.
* Unit testing of formulas is not possible
* Existing models are astoundingly complex and extensively use Excel formulas - traditionally all of which would need to be re-written before the model could be useful in any other computer language (plus key-person risk).
* Companies can't easily communicate the nuances of models (eg; mathematical, technical and financial) when they are expressed in code.
* Managers and domain experts can’t necessarily be expected to code well enough to determine if a model (eg; mathematical, technical and financial) has been translated correctly.
* Data analysis with interesting data sets (large or time series) is hard. Excel can compound this just as your data set becomes interesting.
* Data migrations often require especially skilled programmers, who need to be trained up in the knowledge domain, even if the operation isn't technically difficult.
* Not everyone is going to learn to code, nor should they be expected to.
* People are usually skilled enough in MS Excel but not necessarily in an adequate coding language.
* Scenario analysis usually requires large overhead and can be diﬃcult to manage.
* Data analysts with a strong coding background will do everything they can to express things with Pandas.
* Data analysts who aren't strong coders can do incredible things with Excel but may be causing performance problems and key person risk.


# Features of FlyingKoala

* Provides the ability to unit test Excel formulas using Excel as the calculation engine or Python.
* Elegantly brings together, highlights, and makes available the positive attributes of xlwings, Koala2, Pandas and a number of other libraries without getting in the way.
* Supplies pre-made User Defined Functions for mathematical equations, external application APIs, Python modules and database connectivity.
* Manages caching of models (eg; mathematical, technical, financial, etc...) reducing loading time and takes advantage of a Koala2 feature where an equation can be in workbooks other than the active one.


# Benefits of FlyingKoala

* Can unit test Excel formulas
* Facilitates and encourages domain experts to define a language for their domain and then use the fresh language as the basis for defining models, equations and data related operations where that language can be both processed efficiently by computers and easily understood by other humans.
* Audits are easier because more people know how to read and change Excel formulas than a computer coding language.
* Provides Excel users access to calculation efficiencies which are usually completely unable to access without coding.
* Enables piecemeal migration of existing Excel defined models. eg; Don't _need_ to re-write the entire macro library before making progress on efficient calculation.
* Has potential to reduce key-person risk on pre-existing complex Excel based models
* Considerably reduces the need for a coder to become involved in model development;
  * reduces time for model turnaround,
  * minimizes translation errors,
  * keeps coders in the coding domain,
  * increases re-use of the code written by coders (a single UDF is usually an industry-wide definition).
* The entire mathematical or technical model is available for managers to read because it’s an Excel equation.
* Inter-company and intra-company communication of calculations is considerably improved;
  * all parties no longer require evenly skilled coders,
  * more domain experts can easily read the formulas.
* Makes big data calculations in Excel quicker.
* Multiple mathematical models can be defined and assessed quickly. Great for scenario analysis.
* Extends xlwings to be even more powerful in;
  * Applying Excel formulas to datasets without writing much Python code (in the case of the FlyingKoala UFDs, if any)
  * Data analysis
    * Pre-wrapping some of Pandas classic operations
  * Modelling
    * Financial
    * Mathematical
    * Technical
    * Efficiently evaluating Excel formula calculations by web request (REST) so that the definition of an equation can remain obscured from the domain expert triggering the calculation (eg; a proprietary calculation doesn't leave the premises)
  * System integrations where processes are;
    * Ad-hoc
    * Regular ones where a human needs to audit data
  * More accessible with database connectivity
  * Easier access to functionality found in commonly used Python libraries (Pandas, PVLib, Harvest) by way of pre-defined UDFs

# readthedocs
[The latest documentation](https://flyingkoala.readthedocs.io/en/latest/)

# Examples
These are code examples for using FlyingKoala with the supplied UDFs. For a worked example on how to take advantage of the Koala2 Excel formula reading, read the worked example in the [Introduction PDF](https://github.com/bradbase/flyingkoala/blob/master/doc/Introduction_Article.pdf). The worked example uses the horticulture library to demonstrate the advantages of Koala2 when used in conjunction with the xlwings UDF functionality.

## Horticulture library
There is a library of horticulture related UDFs which assist in calculating Growing Degree-Days. The extent of the Python code you would need to start using the Excel User Defined Function =DegreeDay():

```Python
import xlwings as xw
from flyingkoala import flyingkoala
from flyingkoala.horticulture import *
```

If we were in need of using the pre-defined UDFs which wrap the Pandas resample and other time series functionality:

```Python
import xlwings as xw
from flyingkoala import flyingkoala
from flyingkoala.horticulture import *
from flyingkoala.series import *
```

# TODO:
- [x] Unit testing Excel formulas
- [ ] Change intro document - bring TL;DR into line with README.md
- [x] Release a beta
- [-] Write a UDF which is a generic use case for Koala (eg; takes a variable number of term arguments) **Can't be done**
- [X] Write an Excel plug-in which uses the xlwings REST interface to manage the model cache, and provides the supplied FlyingKoala VBA macros
- [X] Write doco on how to install the add-in by hand
- [ ] Write a script to install the add-in
- [ ] Improve add-in. Requires better handling of essentially everything.
- [ ] Support add-in feature to unload a specific model
- [ ] Write a wizard, to launch from the Excel plug-in, which writes and updates the xlwings Python "code"(/imports) for FlyingKoala defined UDFs
- [ ] Write tests
- [ ] Refactor the timeseries Pandas wrappers
- [ ] Run the accounting code for Harvest
- [ ] Write a MySQL module which behaves in a similar way to the sql extension of xlwings
- [ ] Write a function that queries MySQL and returns results as a dynamic array that fit a worksheet
- [ ] Write a PostgreSQL module which behaves in a similar way to the sql extension of xlwings
- [ ] Write a function that queries PostgreSQL and returns results as a dynamic array that fit a worksheet
- [ ] Write more worked examples showcasing the various FlyingKoala defined functions (both RunPython and UDF)
- [ ] Write a module for [scraping-ebay](https://github.com/cpatrickalves/scraping-ebay)
- [ ] Write a module for an optimization problem using pyomo.
