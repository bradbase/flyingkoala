.. _use:

Using FlyingKoala
=================

This guide assumes you have FlyingKoala, the FlyingKoala add-in and xlwings already installed. If that's not the case, head over to :ref:`installation`.

1. The FlyingKoala Add-In
-------------------------

The FlyingKoala Add-In assists users to manage the cache which holds "models" (/equation systems). This is what it looks like.

.. image:: /images/add-in.PNG
  :alt: Add-in menu

Excel file name
^^^^^^^^^^^^^^^

This is the name of the workbook file. This setting is expected to auto-fill.

In the future we expect to be able to load models from other Excel files which is why this configuration is here.

.. image:: /images/add-in_excel_file_name.PNG
  :width: 400
  :alt: Add-in menu excel file name


Ignore Sheets
^^^^^^^^^^^^^

These are worksheets you would like to not load into the koala cache. Most likely used for worksheets which have your raw data, or worksheets which are not currently being used in the modelling you're working on.

.. image:: /images/add-in_ignore_sheets.PNG
  :width: 400
  :alt: Add-in menu ignore sheets

Reload Koala
^^^^^^^^^^^^

FlyingKoala caches parts of a specified spreadsheet (eg; the one named in "Excel config name"). To extract the part you are interested in, first Koala needs to have a look at the spreadsheet you are interested in. This button does that initial loading. If the spreadsheet has been loaded, this button will load it again. Great for when a formula in the spreadsheet has changed.

.. image:: /images/add-in_reload_koala.PNG
  :width: 75
  :alt: Add-in button to load the specified spreadsheet.

Clear Model Cache
^^^^^^^^^^^^^^^^^

This button clears the FlyingKoala cache of all the loaded models. Handy if you want to re-load some of the equations.

.. image:: /images/add-in_clear_model_cache.PNG
  :width: 75
  :alt: Add-in button to load the specified spreadsheet.

Get Cached Model Names
^^^^^^^^^^^^^^^^^^^^^^

A message box will appear naming all the models which have been loaded into the FlyingKoala model cache.

.. image:: /images/add-in_get_cached_model_names.PNG
  :width: 75
  :alt: Add-in button to display the names of the models which are currently in the FlyingKoala cache.

Get workbook Names
^^^^^^^^^^^^^^^^^^

A message box will appear naming all the named ranges xlwings has access to in the loaded workbook. This helps you check spelling on some of the named ranges. This list can also be seen, and managed, in the Formulas ==> Name Manager menu of Excel.

.. image:: /images/add-in_get_workbook_names.PNG
  :width: 75
  :alt: Add-in button to display the named ranges available to xlwings.

2. The FlyingKoala Configuration
--------------------------------

The FlyingKoala configuration worksheet assists users to manage the relationship between koala and xlwings. It must be named FlyingKoala.conf and this is what it looks like.

.. image:: /images/conf.PNG
  :alt: Add-in menu and configuration worksheet

Currently the config management uses fixed cell references, so don't move anything. (**TODO: open for contribution. There's an example in xlwings for inspiration.**)

This is not kept in an external file as the FlyingKoala operations are generally workbook specific. You are likely to want to have this workbook behave in a particular way -- especially when someone opens the file and re-calc while they are using it.

Excel file name
^^^^^^^^^^^^^^^

This is the name of the workbook file. This setting is expected to auto-fill, but will also be over-written by whatever gets put in the corresponding field in the Add-In.

Ignore Sheets
^^^^^^^^^^^^^

These are a comma delimited list of worksheets you want to have Koala ignore when it loads your spreadsheet into cache. If there is a space in the worksheet name you'll need to use double quotes.

Auto load Koala
^^^^^^^^^^^^^^^

This will allow or deny xlwings the ability to load the workbook when you load UDFs. Basically, it's telling xlwings that when the Python interpreter service starts it's allowed to initialise Koala or not.

* TRUE: the spreadsheet will be loaded into Koala when xlwings interpreter service starts.
* FALSE: The spreadsheet will _not_ be loaded into Koala when xlwings interpreter service starts.

3. Using the FlyingKoala User Defined Functions
-----------------------------------------------
Providing the User Defined Functions (UDFs) you are expecting to use already exist in the FlyingKoala library you can simply import them. They won't be loaded at this point, but will become available for use in Excel like any other formula. An equation will become loaded into Koala cache as a model when you use it.

Make sure you have the dependencies installed for the FlyingKoala module you want to use. Notes on this can be found in :ref:`installation`.

For calculating Growing Degree-Days you would have a module that would look like this:

.. code-block:: Python

  import xlwings as xw
  from flyingkoala import flyingkoala
  from flyingkoala.horticulture import *

For doing some time series transformation you would have a module that would look like this:

.. code-block:: Python

  import xlwings as xw
  from flyingkoala import flyingkoala
  from flyingkoala.timeseries import *

For doing some time series transformation while calculating Growing Degrees-Day you would have a module that would look like this:

.. code-block:: Python

  import xlwings as xw
  from flyingkoala import flyingkoala
  from flyingkoala.horticulture import *
  from flyingkoala.timeseries import *

4. Using the FlyingKoala VBA macros
-----------------------------------
Providing the macro functions you are expecting to use already exist in the FlyingKoala library they will be installed with FlyingKoala.

You'll want to be familiar with writing VBA and reading API style documentation for this one.

Once everything is installed correctly you can call the FlyingKoala VBA functions as you normally do with VBA.


5. Freestyling with FlyingKoala
--------------------------------
If the functionality you want isn't yet supported in a FlyingKoala module, you'll need to write your own (and maybe put it forward to be included in FlyingKoala :D).

This is the advanced approach. It has a rather steep curve. Once you 'get it' things aren't terrible, but I do admit there are a lot of moving parts. The worked example in the example document Introduction_Article.PDF is a great resource.

From here on you'll want to be an particularly familiar with;

* writing Python
* writing xlwings UDFs
* using Excel named ranges, named cells and Excel's Manage Ranges feature

Strap in... Here we go!

At the very core FlyingKoala offers an xlwings friendly interface to Koala2.

Koala2 is a project which can read an MS Excel equation, convert it into Python and then evaluate (run) the Python to produce a result for the equation. FlyingKoala caches the Python code generated for each MS Excel equation and so we can change out the definition of an equation on each call to a function.

To take advantage of the FlyingKoala interface to Koala2 we need to write a function which takes at least two types of argument:

* named range name
* terms for an equation

The named range name is simply a string which identifies the name of a named range containing an equation you want interpreted into Python. This is what the equation will have been keyed on in the cache.

Terms for an equation are variables which will be used while evaluating an equation.

Let us work through writing a Python function taking advantage of FlyingKoala's interface with Koala2 using Growing Degrees Day as the example.

Growing Degrees Day calculations are often specific to a particular application. Potentially differing on region, variety of plants or a number of other influences. When we look at Wikipedia we can find no less than two different equations.

.. math::

  GDD  = \frac{T_{\textrm{max}} + T_{\textrm{min}}}{2} - T_{\textrm{base}}

and

.. math::

  GDD  =  \textrm{max}\left( \frac{T_{\textrm{max}} + T_{\textrm{min}}}{2} - T_{\textrm{base}} \, ,\,0\right)

Consider a situation where we want to do some modelling of those two documented examples and, maybe, develop another method to optimize for our specific situation. It would be great to simply write the equations in Excel while we are running scenarios and developing the new one.

We can see in the above that even though the relationship between the terms is quite different, calculating a Growing Degree Day appears to require the same number of inputs, namely T_min, T_max and T_base. The rest of the relationship is in operators and hard constants. This is an observation which often holds up while developing equations... You only have, or need, particular inputs to find the answer to something and if there is another pathway to the answer it is likely to still only need the same inputs (often enough they are the only inputs you *can get*).

The truly variable terms in the Growing Degree Day equations are T_min and T_max. To obviate some Wikipedia reading, T_base is usually set to a value of 10. This makes T_base a soft constant. One you *might* change but are likely not to. We want to support an ability to change it. Remember, we are trying to invent new mathematics so if there's a knob - we need to be able to turn it.

With that understanding, we can now write a Python function which takes a key for the cache (a named range name) and two parameters being T_min and T_max:

.. code-block:: Python

  import numpy as np
  import pandas as pd

  from flyingkoala import flyingkoala as fk

  def DegreeDay(model, T_min, T_max):
      """Function for calculating Degree Day"""

      inputs_for_DegreeDay = pd.DataFrame({'T_min': np.array([T_min]), 'T_max': np.array([T_max])})
      return fk.EvaluateKoalaModel(model.name.name, inputs_for_DegreeDay)

This function is not finished. It still needs xlwings mark-up to become a User Defined Function. But we can see here that there is nothing genuinely complex about taking three arguments, packaging two into a Pandas Dataframe and then calling EvaluateKoalaModel.

If we put all the xlwings markup on the above Python function, we can import it as a User Defined Function:

.. code-block:: Python

  @xw.func
  @xw.arg('model', xw.Range, doc='Name, as a string, of the model which will be evaluated. The Excel cell name / named range')
  @xw.arg('T_min', np.array, doc='Daily minimum temperature')
  @xw.arg('T_max', np.array, doc='Daily maximum temperature')
  @xw.ret(index=False, header=False)
  def DegreeDay(model, T_min, T_max):
      """Function to assemble a dataframe for calculating Degree Day"""

      if not fk.isKoalaModelCached(model.name.name):
          fk.generateModelGraph(model)

      inputs_for_DegreeDay = pd.DataFrame({'T_min': np.array([T_min]), 'T_max': np.array([T_max])})
      return fk.EvaluateKoalaModel(model.name.name, inputs_for_DegreeDay)

This is essentially all a developer or appropriately skilled data analyst needs to do. The rest is up to the domain expert as it is a case of setting Excel up correctly and then defining the mathematical relationship.

**NOTE**: The mathematical relationship for the equation is not expressed in Python yet. The definition is the responsibility of the domain expert to define in an Excel formula and Koala2 to manage running that definition in Python.

Although setting Excel up is demonstrated in the worked example in the example document Introduction_Article.PDF I'll run through it briefly here.

To set Excel up...

Define named cells for each of the terms. The keys in the anonymous dict which is used to create inputs_for_DegreeDay need to be the same as the names of the named ranges (cell names) that will be used in the Excel formula.

To be specific there needs to be an Excel cell named 'T_min', another called 'T_max'. These names need to be global (can be identified in the Manage Names menu of Excel). These cells need to be referenced in the Excel equation - **NOT** the cell address.

Now we can write an MS Excel formula which will define the relationship between T_min, T_max and T_base. This is the Excel formula for the first GDD equation and is in a cell called Equation_1:

  =((T_max+T_min)/2)-T_base

NOTE: we have used the cell names, not the cell address eg; T_max **not** GDD_formula!B5

Providing the DegreeDay UDF definition is defined, we can use it in Excel:

  =DegreeDay(Equation_1, 3, 25)

That Excel equation will grab the equation from the call called Equation_1 which is =((T_max+T_min)/2)-T_base, convert the equation to Python, set T_min to 3, T_max to 25, T_base to 10, evaluate the result and return it as a value to the cell.
