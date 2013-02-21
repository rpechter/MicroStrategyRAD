MicroStrategyRAD
================

MicroStrategy R Analytic Deployer

Deploying R Analytics to MicroStrategy made easy with the MicroStrategy R Analytic Deployer.

The ability to deploy R Analytics to MicroStrategy combines the Best in Business Intelligence with
the world's fastest growing statistical workbench.  Conceptually, think of the R Analytic as
a Black Box with an R Script inside.  MicroStrategy doesn't need to know much about the R Script in the 
black box, it only needs to know how to pass data in and, after the script 
has executed in the R environment, how to consume the results out.  

The key for MicroStrategy to execute an R Analytic implemented as an R Script is capturing the R 
Analytic's \emph{"signature"}, a description of all inputs and outputs including their number, order and 
nature (data type, such as numeric or string, and size, scalar or vector).  

The MicroStrategy R Analytic Deployer (MicroStrategyRAD) analyses an R Script, identifies all potential variables and 
allows the user to specify the Analytic's signature, including the ability to configure additional
features such as the locations of the script and working directory as well as controlling settings
such as how nulls are handled.  

All this information is added to a header block at the top of the R Script.  The header block, comprised
mostly of non-executing comment lines which are used by MicroStrategy when executing the script.  The 
analytic's signature represented in the header block tells MicroStrategy precisely how to execute the 
R Analytic.

Finally, in order to deploy the R analytic to MicroStrategy, the MicroStrategyRAD provides the metric expression of
each potential output.  The metric expression can be pasted into any MicroStrategy metric editor for 
deploying the R Analytic to MicroStrategy for execution.
