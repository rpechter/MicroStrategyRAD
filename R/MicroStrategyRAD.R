#' Deploying R Analytics to MicroStrategy made easy with the MicroStrategy R Analytic Deployer.
#'
#' The ability to deploy R Analytics to MicroStrategy combines the Best in Business Intelligence with
#' the world's fastest growing statistical workbench.  Conceptually, think of the R Analytic as
#' a Black Box with an R Script inside.  MicroStrategy doesn't need to know much about the R Script in the 
#' black box, it only needs to know how to pass data in and, after the script 
#' has executed in the R environment, how to consume the results out.  
#' 
#' The key for MicroStrategy to execute an R Analytic implemented as an R Script is capturing the R 
#' Analytic's \emph{"signature"}, a description of all inputs and outputs including their number, order and 
#' nature (data type, such as numeric or string, and size, scalar or vector).  
#' 
#' The MicroStrategy R Analytic Deployer (MicroStrategyRAD) analyses an R Script, identifies all potential variables and 
#' allows the user to specify the Analytic's signature, including the ability to configure additional
#' features such as the locations of the script and working directory as well as controlling settings
#' such as how nulls are handled.  
#' 
#' All this information is added to a header block at the top of the R Script.  The header block, comprised
#' mostly of non-executing comment lines which are used by MicroStrategy when executing the script.  The 
#' analytic's signature represented in the header block tells MicroStrategy precisely how to execute the 
#' R Analytic.
#' 
#' Finally, in order to deploy the R analytic to MicroStrategy, the MicroStrategyRAD provides the metric expression of
#' each potential output.  The metric expression can be pasted into any MicroStrategy metric editor for 
#' deploying the R Analytic to MicroStrategy for execution.
#'
#' @references MicroStrategy R Integration Pack. \url{http://MRIP.codeplex.com/}.
#' @docType package
#' @name MicroStrategyRADpackage

#For details on gWidgets, see:
#edit(vignette("gWidgets"))
  
#Load Requred Packages
#  install.packages("gWidgetsRGtk2", dep=TRUE)
  
  require(parser)             #Require the Parser Package
  require(RGtk2)              #Require the RGtk2 Package
  require(gWidgets)           #Require the gWidgets Package
  require(gWidgetsRGtk2)      #Require the gWidgetstRGtk2 Package

  options(guiToolkit = "RGtk2")
  
  DEBUG <- FALSE              #Global Flag for debugging (intended to enable 'print' messages during execution)
  FILE_CHOOSE <- TRUE         #TRUE = Users chooses file; FALSE = Static file name is used
  STICKY_MODS <- TRUE         #FALSE = Controls change when variable changes; TRUE = Controls don't change unless the user changes them
  SPACING <- 10               #Global GUI Spacing value
  MSTR_XEQ_STATEMENT    <- "microstrategy.executionFlag <- (exists(\"mstrExFlag\"))"
  MSTR_XEQ_COMMENT      <- "#Flag if executed by MicroStrategy"
  WORKING_DIR_STATEMENT <- "if(exists(\"workingDir\")) setwd(workingDir)"
  WORKING_DIR_COMMENT   <- "#Working Directory if executed by MicroStrategy"
  TAG_METRIC_EXP        <- "  #Metric Expression: "
  SORT_BY_DEFAULT       <- "{Default = First Input}"

  ### Find index of of substring in string
  Instr <- function(x, s, start=1) {
    #if(DEBUG) print(paste0("Function enter: Instr=", x, " in ", s, " at ", start))
    i <- start                                          #begin with the start character
    slen <- nchar(s)-1                                  #match window width is i+slen
    repeat {
      print(paste0(x, ", ", s, ", ", i, ", ", substr(x, i, i+slen), ", ", (substr(x, i, i+slen)==s)))
      
      bmatch <- (substr(x, i, i+slen)==s)               #did we find a match?
      i <- i+1                                          #go to next character
     if((bmatch)||(i+slen)>nchar(x)) {break}            #repeat until we find a match or reach end of x
    }
    return(ifelse(bmatch, i-1, -1))                     #return index of substring, or -1 if there was no match
  } #End-Instr
  
  ###Track when a parameter is used or un-used
  TrackParams <- function(pName="", isUsed=TRUE) {
    if(DEBUG) print(paste0("Function enter: TrackParams = \"", pName, "\" <- ",isUsed))
    if(nchar(pName)>0) {
      pIndex <- (match(pName, envRAD$paramList))      #Get the index of this parameter name, if any
      if(is.na(pIndex)) {                             #If that's not a valid parameter, return FALSE
        return(FALSE)
      } else {  
        envRAD$paramUsed[pName] <- isUsed
        if(envRAD$allowUpdates) {                     #Update UI, but only if Updates are allowed
if(DEBUG) print("A")        
          envRAD$lstParam[] <- envRAD$paramList[!envRAD$paramUsed]
          update(envRAD$lstParam)
if(DEBUG) print("B")
        }
      }
    }
    return(TRUE)
  }
  
  ### Get the parameter from a MicroStrategy variable comment line
  GetParam <- function(x) {
    if(DEBUG) print(paste0("Function enter: GetParam=", x))
    pName <- ""                                         #Default result to return
    tokens <- unlist(strsplit(x,"[[:space:]]+"))        #Break up string into tokens:  For a parameter there should only be three, in the form "varName -p/-parameter pName"
    if(length(tokens)>=3) {                             #Must have at least 3 tokens for a parameter
      if(tolower(substr(tokens[2],1,2))=="-p") {        #If the second token starts with "-p"
        pName <- tokens[3]                              #pName is the third token
        if(!TrackParams(pName, TRUE)) {                 #If that's not a valid parameter, show dialog
          wParam <- gbasicdialog("Invalid Parameter", horizontal=FALSE, handler=function(h, ...){
            pName <<- svalue(lParam)                    #User has selected to replace with an existing parameter
            TrackParams(pName, TRUE)                    #Track that this parameter is used
          })
          glabel(paste0("This invalid parameter was found in the script."), container=wParam, anchor=c(-1,0))
          glabel(paste0("\"", pName, "\""), container=wParam, anchor=c(-1,0))
          glabel(paste0("Please select the correct parameter for it:"), container=wParam, anchor=c(-1,0))
          lParam <- gdroplist(envRAD$paramList[!envRAD$paramUsed], container=wParam, editable=FALSE, selected=1)
          visible(wParam, set=TRUE) # show dialog
        }
      }
    }
    return(pName)
  } #End-GetParam
  
  ### Set Function Type based on the nature of inputs and outputs
  SetFunctionID <- function(...) {    
    if(DEBUG) print("Function enter: SetFunctionID")
    inputCount <- length(envRAD$g.tRAD[((envRAD$g.tRAD[, 2]=="Input")), 1])        #Count the number of inputs
    inputVectorCount <- length(
                          envRAD$g.tRAD[((envRAD$g.tRAD[, 2]=="Input")
                                         &(envRAD$g.tRAD[, 4]=="Vector")), 1])     #Count the number of inputs that are vectors
    bVectorInput <- (inputVectorCount>0)                                           #Flag is any inputs are vectors
    outputCount <- length(envRAD$g.tRAD[((envRAD$g.tRAD[, 2]=="Output")), 1])      #Count the number of outputs
    outputVectorCount <- length(
                          envRAD$g.tRAD[((envRAD$g.tRAD[, 2]=="Output")
                                         &(envRAD$g.tRAD[, 4]=="Vector")), 1])     #Count the number of outputs that are vectors
    bVectorOutput <- (outputVectorCount>0)                                         #DEPRECATED: Old way where all outputs had to be either vector or Scalar
    if ((inputCount==0)||(outputCount==0)) {                                       #Make sure we have at least one input and at least one output
      envRAD$functionID <- 7                                                       #Need to have at least one input or output, ERROR-7
    } else {                                                                       #Ok, We have at least one input and at least one output
#      if (!((outputVectorCount==outputCount)||(outputVectorCount==0))) {           #DEPRECATED: Make sure all outputs are the same type
#        envRAD$functionID <- 9                                                     #DEPRECATED: Outputs must all be either all scalar or all vector, ERROR-9     
#      }  else {                                                                    #DEPRECATED: Ok, set the functionID using this Truth Table
      outputParamTypes <- envRAD$g.tRAD[envRAD$g.tRAD[, 2]=="Output", 4]           #Get the parameter types of all the output variables
      bVectorOutput <- (outputParamTypes[svalue(envRAD$lstOut, 
                                                index=TRUE)]=="Vector")            #Flag if this output is a vector
        envRAD$functionID <- ifelse(!bVectorInput&&!bVectorOutput, 1,              # VectorInputs VectorOutput  functionID
                                ifelse(bVectorInput&&bVectorOutput, 2,             #      0           FALSE     Simple-1           
                                       ifelse(bVectorInput&&!bVectorOutput,3, 6))) #     >=1          TRUE      Relative-2
#      }                                                                           #     >=1          FALSE     Agg-3
    }                                                                              #      0           TRUE      ERROR-6
    if((envRAD$functionID<2)||(envRAD$functionID>3)) {                             #Handle SortBy Option based on functionID
      enabled(envRAD$chkSort) <- FALSE                                             #Disable SortBy Option if not Relative-2 or Agg-3
    } else {                                                                       #functionID must be Relative-2 or Agg-3
      enabled(envRAD$chkSort) <- TRUE                                              #Enable SortBy 0ption since functionID is Relative-2 or Agg-3 
      if(svalue(envRAD$chkSort)) {                                                 #Determine if SortBy Option is Selected
        inputsParamType <- envRAD$g.tRAD[(envRAD$g.tRAD[, 2]=="Input"), 4]         #SortBy is selected, get scalar/vector type of the first input 
        if(inputsParamType[1]=="Vector") {                                         #Determine if the first input is a vector (a requirement for SortBy)
          envRAD$functionID <- envRAD$functionID + 2                               #First input is a vector so use SortBy Function: Relative-2 --> RelativeS-4, Agg-3 --> AggS-5
        } else {                                                                   #Otherwise, first input is not a vector but SortBy is selected
          envRAD$functionID <- 8                                                   #Need to have vector first input to select Sort By --> ERROR-8
        }
      }
    }
    svalue(envRAD$eFuncType) <- envRAD$funcList[envRAD$functionID]
    svalue(envRAD$tFuncDesc) <- envRAD$funcDesc[envRAD$functionID]
    svalue(envRAD$tInputs)   <- paste0(inputCount, " input", GetVarInfo(inputCount, inputVectorCount))
    svalue(envRAD$tOutputs)  <- paste0(outputCount, " output", GetVarInfo(outputCount, outputVectorCount))
  }  #End-SetFunctionID
  
  ### Concatenate two strings with a comma, if the first string is not empty
  CommaCat <- function(A, B) {
    return(paste0(A, ifelse((nchar(A)>0), ", ", ""), B))
  }  #End-CommaCat
  
  ### Get string for variables
  GetVarInfo <- function(totalCount, vectorCount) {
    return(paste0(ifelse(totalCount==1,
                         ifelse(vectorCount==0," (scalar)", " (vector)"),
                         ifelse(vectorCount==0, "s (all scalar)", 
                                ifelse(totalCount==vectorCount, "s (all vector)", 
                                       paste0("s (",vectorCount, " vector, ", totalCount-vectorCount, " scalar)"))))))
  }
  
  ### Set the available outputs, if any
  SetOutputs <- function(...) {
    if(!envRAD$allowUpdates) return()                                    #Don't process actions while updates are NOT allowed
    if(DEBUG) print("Function enter: SetOutputs")    
    envRAD$allowUpdates <- FALSE                                         #Temporarily block updates
    selOut <- svalue(envRAD$lstOut)                                      #Remember the currently selected output
    envRAD$lstOut[] <- envRAD$g.tRAD[envRAD$g.tRAD[, 2]=="Output", 1]    #Update the gdroplist control with the current outputs
    update(envRAD$lstOut)                                                #Update the control
    if(length(envRAD$lstOut)==0) {                                       #Check if there are no outputs
      enabled(envRAD$lstOut) <- FALSE                                    #When there are no outputs, disable the gdroplist control 
      envRAD$lstOut[] <- c("<-- No Outputs! -->")                        #Add one output so the user will know that there's no outputs
      svalue(envRAD$lstOut, index=TRUE) <- 1                             #Select the first (and only) output in the list
    } else {                                                             #There's at least one output
      enabled(envRAD$lstOut) <- TRUE                                     #Make sure the gdroplist control is enabled when there's at least one output
      envRAD$lstOut[1] <- paste0(envRAD$lstOut[1]," (default)")          #Indicate that the first output is the default
      if(is.null(selOut)) {                                              #If there was no selected output to remember
        svalue(envRAD$lstOut, index=TRUE) <- 1                           #Then select the first output
      } else {                                                           #Otherwise, select the remembered output or, if that one is no longer in the list, select the first output
        svalue(envRAD$lstOut, index=TRUE) <- match(selOut, envRAD$lstOut[], nomatch=1)
      }
    }    
    envRAD$allowUpdates <- TRUE                                          #Re-enable updates
  }  #End-SetOutputs
  
  ### Create/Update the Metric Expression
  UpdateExpression <- function(thisOutput) {
    if(!envRAD$allowUpdates) return()    #Don't process actions while updates are NOT allowed
    if(DEBUG) print("Function enter: UpdateExpression")    
    SetFunctionID()  #Set the function type
    if (envRAD$functionID>5) {
      svalue(envRAD$txtMetricExp) <- paste0("<-- Function Error: ", envRAD$funcDesc[envRAD$functionID], " -->")
    } else {
      #Set the Function Parameters 
      sParams <- paste0("[_RScriptFile]=\"", svalue(envRAD$edtScr), "\"")
      if(!svalue(envRAD$chkCIC))              sParams <- CommaCat(sParams, "[_CheckInputCount]=False")
      if(!svalue(envRAD$chkNul))              sParams <- CommaCat(sParams, "[_NullsAllowed]=False")
      if(is.na(thisOutput)) {
        if (!is.null(svalue(envRAD$lstOut))) {
          if(svalue(envRAD$lstOut, index=TRUE)>1) sParams <- CommaCat(sParams, paste0("[_OutputVar]=\"", svalue(envRAD$lstOut), "\""))
        }
      } else {
        if(thisOutput>1) sParams <- CommaCat(sParams, paste0("[_OutputVar]=\"", envRAD$lstOut[thisOutput], "\""))        
      }
      if(svalue(envRAD$chkDir)) sParams <- CommaCat(sParams, paste0("[_WorkingDir]=\"", svalue(envRAD$edtDir), "\""))
      if(svalue(envRAD$chkSort)&&(svalue(envRAD$edtSrt)!=SORT_BY_DEFAULT)&&(nchar(svalue(envRAD$edtSrt))>0)) { 
        sParams <- CommaCat(sParams, paste0("SortBy=(", svalue(envRAD$edtSrt), ")"))
      }
      if(nchar(sParams)>0)      sParams <- paste0("<", sParams, ">")
      
      #Set the Metric Arguments
      Args <- envRAD$g.tRAD[envRAD$g.tRAD[, 2]=="Input", 1]
      sArgs <- Args[1]
      if(length(Args)>1) {
        for(i in 2:length(Args)) sArgs <- paste0(sArgs, ", ", Args[i])
      } 
      if(is.na(thisOutput)) {
        svalue(envRAD$txtMetricExp) <- paste0(envRAD$funcName[envRAD$functionID], sParams, paste0("(", sArgs, ")"))      
      } else {
        return(paste0(envRAD$funcName[envRAD$functionID], sParams, paste0("(", sArgs, ")"))) 
      }
    }
    
  } #End-UpdateExpression
  
  ### Function to save the R Script with the Function Signature comment block at the top
  ScriptSave <- function(...) {
    if(DEBUG) print("Function enter: ScriptSave")
    if(envRAD$functionID > 5)  {
      gmessage(paste0("Please fix this error before saving:\n\n", envRAD$funcDesc[envRAD$functionID]), title="Cannot Save Due to Error", icon="error")
      return()
    }
    #Create the comment block
    commentBlock <- ""
    numInputs <- sum(envRAD$g.tRAD[, 2]=="Input")
    rptCount <- svalue(envRAD$RptCt)
    inputCount <- 0
    outputCount <- 0
    for (i in 1:nrow(envRAD$g.tRAD)) {
      if(envRAD$g.tRAD[i, 2]=="Parameter") {
        commentBlock <- paste0(commentBlock, "#", envRAD$g.tRAD[i, 1], " -parameter ", envRAD$g.tRAD[i, 5], "\n")        
      } else {
        if(envRAD$g.tRAD[i, 2]=="Input") {
          inputCount <- inputCount + 1
          if(inputCount>(numInputs-rptCount)) sRepeat=" -repeat" else sRepeat=""
        } else {
          if(envRAD$g.tRAD[i, 2]=="Output") {
            outputCount <- outputCount + 1
            sRepeat <- paste0(TAG_METRIC_EXP, UpdateExpression(outputCount))
          } else sRepeat="" 
        }
        commentBlock <- paste0(commentBlock, "#", envRAD$g.tRAD[i, 1], " -", tolower(envRAD$g.tRAD[i, 2]), 
                               ifelse((envRAD$g.tRAD[i, 3]=="Default"), "", paste0(" -", tolower(envRAD$g.tRAD[i, 3]))), 
                               " -", tolower(envRAD$g.tRAD[i, 4]), sRepeat, "\n")
      }
    } 
    
    commentBlock <- paste0(commentBlock, MSTR_XEQ_STATEMENT, "  ", MSTR_XEQ_COMMENT, "\n")    
    if(svalue(envRAD$chkDir)) commentBlock <- paste0(commentBlock, WORKING_DIR_STATEMENT, "  ", WORKING_DIR_COMMENT, "\n")
    commentBlock <- paste0("#MICROSTRATEGY_BEGIN\n", commentBlock, "#MICROSTRATEGY_END\n")

    #Create and show the Save Dialog
    wSave <- gwindow("Save")
    gSave <- ggroup(container = wSave, horizontal=FALSE)
    fCB <- gframe("This header block will be added to the top of your script:", container=gSave, anchor=c(-1,0), expand=TRUE)
    tCB <- gtext(commentBlock, container=fCB, expand=TRUE, font.attr=list(family="monospace"))
    fSc <- gexpandgroup("R Script:", container=gSave, expand=FALSE)
    tSc <- gtext(paste0(envRAD$g.script, collapse="\n"), container=fSc, expand=TRUE, font.attr=list(family="monospace"), height=500)
    #addSpring(gSave)
    gButtons <- ggroup(container = gSave, horizontal=TRUE)    ## A group to organize the buttons
    addSpring(gButtons)                                       ## Push buttons to right
    gbutton("Save", container=gButtons, 
            handler = function(h, ...) {
              sErrMsg <- tryCatch({  
                saveFile <- file(svalue(envRAD$edtScr), "w")
                cat(commentBlock, file=saveFile, sep="\n")
                cat(envRAD$g.script, file=saveFile, sep="\n")
                close(saveFile)
                envRAD$saved <- TRUE
                gmessage("\nScript saved successfully!\n", icon="info", title="Save")
                visible(envRAD$gME) <- TRUE                           #Show the metric expression, now that the script changes have been saved
                scriptErrMsg<-""                                      #If we made it here, then there's no errors to report  
              }, error = function(err) {  
                return(geterrmessage())                               #Report error that was caught
              })
              if(nchar(sErrMsg)>0) {
                gmessage(paste0("The script was not able to be saved due to the following error: ","\n\n", sErrMsg), 
                         title="Script Not Saved", icon="error")
              }
              dispose(wSave)
    })
    gbutton("Cancel", container=gButtons, handler = function(h, ...) dispose(wSave))    
    
  } #End-ScriptSave
  
  ### Function to save the R Script with the Function Signature comment block at the top
  ScriptSave2 <- function(...) {
    if(DEBUG) print("Function enter: ScriptSave2")
    if(envRAD$functionID > 5)  {
      gmessage(paste0("Please fix this error before saving:\n\n", envRAD$funcDesc[envRAD$functionID]), title="Cannot Save Due to Error", icon="error")
      return()
    }
    #Create the comment block
    commentBlock <- ""
    
    #Inputs first:
    rptCount <- svalue(envRAD$RptCt)
    inputCount <- 0
    inList <- grep("Input", envRAD$g.tRAD[,2])
    numInputs <- length(inList)
    for (i in inList) {
      inputCount <- inputCount + 1
      commentBlock <- paste0(commentBlock, "#", envRAD$g.tRAD[i, 1], " -", tolower(envRAD$g.tRAD[i, 2]), 
                             ifelse((envRAD$g.tRAD[i, 3]=="Default"), "", paste0(" -", tolower(envRAD$g.tRAD[i, 3]))), 
                             " -", tolower(envRAD$g.tRAD[i, 4]), 
                             ifelse(inputCount>(numInputs-rptCount)," -repeat", ""), "\n")
    }
    
    outputCount <- 0
    outList <- grep("Output", envRAD$g.tRAD[,2])
    numOutputs <- length(outList)
    commentBlock <- paste0(commentBlock, "#\n")
    for (i in outList) {
      outputCount <- outputCount + 1
      commentBlock <- paste0(commentBlock, "#", envRAD$g.tRAD[i, 1], " -", tolower(envRAD$g.tRAD[i, 2]), 
                             ifelse((envRAD$g.tRAD[i, 3]=="Default"), "", paste0(" -", tolower(envRAD$g.tRAD[i, 3]))), 
                             " -", tolower(envRAD$g.tRAD[i, 4]), 
                             paste0(TAG_METRIC_EXP, UpdateExpression(outputCount)), "\n")
    }
    
    
    pList <- grep("Parameter", envRAD$g.tRAD[,2])
    numParams <- length(pList)
    if(numParams>0) {
      commentBlock <- paste0(commentBlock, "#\n")
      for (i in pList) {
        commentBlock <- paste0(commentBlock, "#", envRAD$g.tRAD[i, 1], " -parameter ", envRAD$g.tRAD[i, 5], "\n")        
      }
    }
    
    dList <- grep("Disabled", envRAD$g.tRAD[,2])
    numDis <- length(dList)
    if(numDis>0) {
      commentBlock <- paste0(commentBlock, "#\n")
      for (i in dList) {
        commentBlock <- paste0(commentBlock, "#", envRAD$g.tRAD[i, 1], " -", tolower(envRAD$g.tRAD[i, 2]), 
                               ifelse((envRAD$g.tRAD[i, 3]=="Default"), "", paste0(" -", tolower(envRAD$g.tRAD[i, 3]))), 
                               " -", tolower(envRAD$g.tRAD[i, 4]), "\n")
      }
    }
    commentBlock <- paste0(commentBlock, "#\n", MSTR_XEQ_STATEMENT, "  ", MSTR_XEQ_COMMENT, "\n")    
    if(svalue(envRAD$chkDir)) commentBlock <- paste0(commentBlock, WORKING_DIR_STATEMENT, "  ", WORKING_DIR_COMMENT, "\n")
    commentBlock <- paste0("#MICROSTRATEGY_BEGIN\n", commentBlock, "#MICROSTRATEGY_END\n")
        
    #Create and show the Save Dialog
    wSave <- gwindow("Save")
    gSave <- ggroup(container = wSave, horizontal=FALSE)
    fCB <- gframe("This header block will be added to the top of your script:", container=gSave, anchor=c(-1,0), expand=TRUE)
    tCB <- gtext(commentBlock, container=fCB, expand=TRUE, font.attr=list(family="monospace"))
    fSc <- gexpandgroup("R Script:", container=gSave, expand=FALSE)
    tSc <- gtext(paste0(envRAD$g.script, collapse="\n"), container=fSc, expand=TRUE, font.attr=list(family="monospace"), height=500)
    #addSpring(gSave)
    gButtons <- ggroup(container = gSave, horizontal=TRUE)    ## A group to organize the buttons
    addSpring(gButtons)                                       ## Push buttons to right
    gbutton("Save", container=gButtons, 
            handler = function(h, ...) {
              sErrMsg <- tryCatch({  
                saveFile <- file(svalue(envRAD$edtScr), "w")
                cat(commentBlock, file=saveFile, sep="\n")
                cat(envRAD$g.script, file=saveFile, sep="\n")
                close(saveFile)
                envRAD$saved <- TRUE
                gmessage("\nScript saved successfully!\n", icon="info", title="Save")
                visible(envRAD$gME) <- TRUE                           #Show the metric expression, now that the script changes have been saved
                scriptErrMsg<-""                                      #If we made it here, then there's no errors to report  
              }, error = function(err) {  
                return(geterrmessage())                               #Report error that was caught
              })
              if(nchar(sErrMsg)>0) {
                gmessage(paste0("The script was not able to be saved due to the following error: ","\n\n", sErrMsg), 
                         title="Script Not Saved", icon="error")
              }
              dispose(wSave)
            })
    gbutton("Cancel", container=gButtons, handler = function(h, ...) dispose(wSave))    
    
  } #End-ScriptSave2

  ### Function for when a variable selection is changed
  VarChanged <- function(h, ...) {
    if(!envRAD$allowUpdates) return()    #Don't process actions while updates are NOT allowed
    if(DEBUG) print(paste0("Function enter: VarChanged ='", svalue(h$obj), "'[", svalue(h$obj, index=TRUE) , "] currentItem=", envRAD$currentItem))
    if(length(svalue(h$obj))==0) {
      #There is no selected variable -- we must have entered this handler after deleting the current variable
      #So, let's select the next variable or, if there's no next variable select the end of the list
      svalue(h$obj, index=TRUE) <- ifelse(envRAD$currentItem > nrow(envRAD$g.tRAD), nrow(envRAD$g.tRAD), envRAD$currentItem)       
      if(DEBUG) print(paste0("----> No current item, force select = ", svalue(h$obj, index=TRUE)))
    }
    envRAD$currentItem <- svalue(h$obj, index=TRUE)
    if(DEBUG) print(paste0("--> CurrentItem changed to = ", envRAD$currentItem))
    enabled(envRAD$gMove) <- TRUE                                                         #Enable moving since a variable has been selected
    if(envRAD$g.tRAD[envRAD$currentItem[1], 2]=="Parameter") {                            #Changing to a Parameter?
      svalue(envRAD$radioDir) <- "Parameter"                                              #  Set Direction radio button control and nothing else
      enabled(envRAD$pValue) <- TRUE
      #svalue(envRAD$lstParam) <- envRAD$g.tRAD[envRAD$currentItem[1], 5] 
      svalue(envRAD$pValue) <- envRAD$g.tRAD[envRAD$currentItem[1], 6]
    } else {
      enabled(envRAD$pValue) <- FALSE
      if(!STICKY_MODS) {                                                           #Otherwise, if "not STICKY" (Change info when vars change)
        svalue(envRAD$radioDir) <- envRAD$g.tRAD[envRAD$currentItem[1], 2]                #Set Direction radio button control
        if(nchar(envRAD$g.tRAD[envRAD$currentItem[1], 3])>0) {                            #If Data Type radio button control is not null
          svalue(envRAD$radioType) <- envRAD$g.tRAD[envRAD$currentItem[1], 3]             #  Then set the Data Type
        }
        if(nchar(envRAD$g.tRAD[envRAD$currentItem[1], 4])>0) {                            #If Parameter Type radio button control is not null
          svalue(envRAD$radioSize) <- envRAD$g.tRAD[envRAD$currentItem[1], 4]             #  Then set the Parameter Type
        }
      }
    }
  } #End-VarChanged
      
  ### Function to move a variable up or down
  VarMove <- function(h, ...) {
    if(DEBUG) print(paste("Function enter: VarMove =", svalue(h$obj)))
    if (envRAD$currentItem>0) {
      if(svalue(h$obj)=="Up") {
        if(envRAD$currentItem[1]>1) {
          for(i in 1:length(envRAD$currentItem)) {
            tmp <- envRAD$g.tRAD[envRAD$currentItem[i], ]
            envRAD$g.tRAD[envRAD$currentItem[i], ] <- envRAD$g.tRAD[envRAD$currentItem[i]-1, ]
            envRAD$g.tRAD[envRAD$currentItem[i]-1, ] <- tmp
          }
          svalue(envRAD$g.tRAD) <- envRAD$currentItem-1
        }
      }
      else {
        if(envRAD$currentItem<nrow(envRAD$g.tRAD)) {
          for(i in length(envRAD$currentItem):1) {
            tmp <- envRAD$g.tRAD[envRAD$currentItem[i], ]
            envRAD$g.tRAD[envRAD$currentItem[i], ] <- envRAD$g.tRAD[envRAD$currentItem[i]+1, ]
            envRAD$g.tRAD[envRAD$currentItem[i]+1, ] <- tmp
          }
          svalue(envRAD$g.tRAD) <- envRAD$currentItem+1
        }
      }
      envRAD$saved <- FALSE  #Clear the saved flag since a change has been made
      SetOutputs()           #Update the outputs, since they could have changed order
      UpdateExpression(NA)   #Update the expression to reflect the changes
    }
  } #End-VarMove
  
  ### Function to apply changes to a variable
  VarApply <- function(...) {
    if(DEBUG) print(paste("Function enter: VarApply: envRAD$currentItem=", envRAD$currentItem, "Type=", svalue(envRAD$radioType), "Size=", svalue(envRAD$radioSize), "Param=", svalue(envRAD$lstParam)))
    if((svalue(envRAD$radioDir)=="Output")&&(svalue(envRAD$radioType)=="Default")) {  #Handle the case where the user is trying to set an output with a default datatype
      gmessage("Outputs do not have a default, please select a specific data type.", title="Specific Data Type Needed", icon="error")
    } else {
      if(length(svalue(envRAD$g.tRAD))>0) envRAD$currentItem <- svalue(envRAD$g.tRAD, index=TRUE)
      if(svalue(envRAD$radioDir)=="Parameter") {                           #If this is a parameter being set, set only the first selected
        envRAD$g.tRAD[envRAD$currentItem[1], 2] <- svalue(envRAD$radioDir) #Set the direction to Parameter
        envRAD$g.tRAD[envRAD$currentItem[1], 3] <- ""                      #Parameters have no specific data type (that's defined by the parameter used)
        envRAD$g.tRAD[envRAD$currentItem[1], 4] <- ""                      #Parameters have no specific parameter type (they're always scalar)
        TrackParams(envRAD$g.tRAD[envRAD$currentItem[1], 5], FALSE)        #Set the previous parameter to un-used
        envRAD$g.tRAD[envRAD$currentItem[1], 5] <- svalue(envRAD$lstParam) #Set the new parameter
        SetParamValue(envRAD$currentItem[1])                               #Set the parameter value
        TrackParams(svalue(envRAD$lstParam), TRUE)                         #Set the new parameter to used
      } else {                                                             #This is NOT a parameter being set, set all selected
        envRAD$g.tRAD[envRAD$currentItem, 2] <- svalue(envRAD$radioDir)    #Set the direction
        envRAD$g.tRAD[envRAD$currentItem, 3] <- svalue(envRAD$radioType)   #Set the data type
        envRAD$g.tRAD[envRAD$currentItem, 4] <- svalue(envRAD$radioSize)   #Set the parameter type
        envRAD$g.tRAD[envRAD$currentItem, 5] <- ""                         #Clear the parameter
      }
      envRAD$saved <- FALSE                                                #Clear the saved flag since a change has been made
      SetOutputs()                                                         #Update the outputs, they could have changed
      UpdateExpression(NA)                                                   #Update the expression to reflect the changes
    }
  } #End-VarApply
  
  ### Function to delete a variable
  VarDelete <- function(...) {
    if(DEBUG) print(paste0("Function enter: varDelete, envRAD$currentItem=", envRAD$currentItem))
    envRAD$allowUpdates <- FALSE                               #Don't allow updates due to the removal in the next statement
    envRAD$deletedVars <- rbind(envRAD$deletedVars, envRAD$g.tRAD[envRAD$currentItem, ])  #Remember deleted vars so they can be un-deleted
    TrackParams(envRAD$g.tRAD[envRAD$currentItem, 5], FALSE)
    envRAD$g.tRAD[, ] <- envRAD$g.tRAD[-envRAD$currentItem, ]  #Remove the currently selected items
    enabled(envRAD$btnUnDel) <- TRUE                           #Enable the UnDelete button
    if(envRAD$currentItem>nrow(envRAD$g.tRAD)) envRAD$currentItem <- nrow(envRAD$g.tRAD) #Handle when last row is deleted -- thanks again Li! 
    envRAD$allowUpdates <- TRUE                                #Re-allow updates again
    envRAD$saved <- FALSE                                      #Clear the saved flag since a change has been made
    SetOutputs()                                               #Update the outputs, one could have been deleted
    UpdateExpression(NA)                                         #Update the expression to reflect the changes
  } #End-VarDelete
  
  ### Function to un-delete the last deleted variable
  VarUnDel <- function(...) {
    if(DEBUG) print(paste0("Function enter: varUnDel, envRAD$currentItem=", envRAD$currentItem))
    envRAD$allowUpdates <- FALSE                               #Don't allow updates due to the addition in the next statement
    envRAD$g.tRAD[] <- rbind(envRAD$g.tRAD[, ], envRAD$deletedVars[nrow(envRAD$deletedVars), ])  #Restore the most recently selected item
    TrackParams(envRAD$deletedVars[nrow(envRAD$deletedVars), 5], TRUE)
    envRAD$deletedVars <- envRAD$deletedVars[-nrow(envRAD$deletedVars), ]  #Remove the most recently deleted var from the list
    if(nrow(envRAD$deletedVars)==envRAD$delVarMin) enabled(envRAD$btnUnDel) <- FALSE
    envRAD$allowUpdates <- TRUE                                #Re-allow updates again
    envRAD$saved <- FALSE                                      #Clear the saved flag since a change has been made
    SetOutputs()                                               #Update the outputs, one could have been deleted
    UpdateExpression(NA)                                         #Update the expression to reflect the changes
  } #End-VarUnDel
  
  SetParamValue <- function(pIndex) {
    if(DEBUG) print(paste0("Function enter: SetParamValue, pIndex=", pIndex))
    if(pIndex>0) {
      if(envRAD$g.tRAD[pIndex, 2]=="Parameter") {
        pTyp <- substr(envRAD$g.tRAD[pIndex, 5], 1, 1)
        if(pTyp=="B") {
          envRAD$g.tRAD[pIndex, 6] <- ifelse(tolower(substr(svalue(envRAD$pValue), 1, 1))=="t",TRUE,FALSE) 
        } else {
          if(pTyp=="N") {
            if(is.numeric(svalue(envRAD$pValue))) {
              envRAD$g.tRAD[pIndex, 6] <- svalue(envRAD$pValue)
            } 
          } else {
            envRAD$g.tRAD[pIndex, 6] <- svalue(envRAD$pValue)
          } 
        } 
      }      
    }
  }  #End-SetParamType
  
  ### Create main dialog
  CreateDialog <- function(dfVar, ...) {
    if(DEBUG) print("Function enter: CreateDialog")
        
    #Top of the Dialong
    gRAD <- ggroup(horizontal=FALSE, spacing=SPACING, container = envRAD$winRAD)      #Group container for main dialog
    visible(gRAD) <- FALSE                                                            #Hide this group while it's being created
    fScrpt <- gframe("R Script File", cont=gRAD, spacing=SPACING, expand=FALSE)             #Frame to hold the Script Name
    envRAD$edtScr <- gedit("", editable=TRUE, cont=fScrpt, expand=TRUE)               #edtScr is the edit box with the Script name
    addHandlerKeystroke(envRAD$edtScr, UpdateExpression(NA))                              #Add handler for editing the script file location
      
    #Upper Panel -- Modify and Function Info
    grpUpper <- ggroup(cont=gRAD, spacing=SPACING)                                    #Group for the upper set of controls (above the variables table, below the Script name)
    frmVars <- gframe("Modify Selected Variables", 
                      cont=grpUpper, expand=FALSE, spacing=SPACING)                   #Frame for Modifying Selected Variables
    layUpper <- glayout(container = frmVars, spacing=SPACING)                         #Layout containter for holding the controls that modify variables
    layUpper[1, 1] <- fDir <- gframe("Direction", cont=layUpper, spacing=SPACING)     #Left most in the Layout is Direction container, which is a frame
    envRAD$radioDir <- gradio(envRAD$dirList, horizontal=FALSE, cont=fDir, 
                        spacing=SPACING, handler=function(h, ...) {                   #Handler for Direction Radio button:
                          if(svalue(h$obj, index=TRUE)==3) {                          #If direction is Parameter, then
                            enabled(envRAD$radioType) <- FALSE                        #  Disable Data Type, a parameter's data type is determined by the type of parameter selected
                            enabled(envRAD$radioSize) <- FALSE                        #  Disable Parameter Type, parameters are always scalar
                            enabled(envRAD$lstParam) <- TRUE                          #  Enable the Parameter so the user can pick one.
                            enabled(envRAD$btnApply) <- (nchar(svalue(envRAD$lstParam))>0)   #  Disable the Apply button if there is no parameter selected 
                          } else {                                                    #Else, the direction is NOT parameter, so
                            enabled(envRAD$radioType) <- TRUE                         #  Enable the Data Type
                            enabled(envRAD$radioSize) <- TRUE                         #  Enable the Parameter Type
                            enabled(envRAD$lstParam) <- FALSE                         #  Disable the Parameter drop down
                            enabled(envRAD$btnApply) <- TRUE                          #  Enable the Apply button
                          }
                        })                                                            #Radio button control for Direction and it's hander
    layUpper[1, 2] <- fType <- gframe("Data Type", cont=layUpper, expand=TRUE)        #2nd from the Left in the Upper Layout is Data Type container, which is a group
    envRAD$radioType <- gradio(envRAD$dtypeList[2:4], cont=fType, spacing=SPACING)    #Radio button control for Data Type (requires no handler)
    layUpper[1, 3] <- gUp3 <- ggroup(cont=layUpper, horizontal=FALSE)                 #3rd from the left in the Upper Layout is the Parameter Type and Repeat Count
    fSize <- gframe("Parameter Type", cont=gUp3, horizontal=FALSE, expand=TRUE)       #Container frame for Parameter type
    envRAD$radioSize <- gradio(envRAD$ptypeList[2:3], cont=fSize, spacing=SPACING)    #Radio button control for Parameter Type (requires no handler)
    addSpring(gUp3)                                                                   #Push the Repeat Count container to the bottom
    fRC <- gframe("Repeat Count", cont=gUp3, spacing=SPACING)                         #Container frame for Repeat Count
    envRAD$RptCt <- gspinbutton(from=0, to=100, by=1, value=0, cont=fRC, expand=TRUE) #Spin button control for the Repeat Count
    layUpper[1, 4] <- gUp4 <- ggroup(cont=layUpper, horizontal=FALSE)                 #4th from the left in the Upper Layout is for Parameter and Actions
    fParam <- gframe("Parameter", cont=gUp4, horizontal=FALSE, spacing=SPACING)       #Container frame for Parameter
    envRAD$lstParam <- gdroplist(envRAD$paramList[!envRAD$paramUsed], editable=FALSE, 
                                 enabled=FALSE, cont=fParam, spacing=SPACING, 
                                 handler=function(...) {                              #Handler for Parameter Dropdown
                                   if(svalue(envRAD$radioDir)=="Parameter") {         #If Directions is Parameter and the selected Parameter is not blank, then enable the apply button; otherwise, disable the apply button
                                     enabled(envRAD$btnApply) <- (nchar(svalue(envRAD$lstParam))>0)
                                   }
                                 })                                                   #Dropdown control for Parameters and it's handler
    gPval <- ggroup(cont=fParam, spacing=SPACING)
    glabel("Value", cont=gPval)
    envRAD$pValue <- gedit("", editable=TRUE, expand=TRUE, cont=gPval, 
          handler=SetParamValue(envRAD$currentItem[1]))                              #Edit box for Parameter Value 
    enabled(envRAD$lstParam) <- FALSE                                                 #Disable the Parameter Dropdown since initially the Direction is NOT set to Parameter
    frmApply <- gframe("Act on Selected Variables", horizontal=TRUE, cont=gUp4, expand=TRUE)  #Container frame for Actions
    
    envRAD$btnApply <- gbutton("Apply", handler=VarApply, cont=frmApply)                     #Apply button
    btnDelete <- gbutton("Delete", handler=VarDelete, cont=frmApply)                  #Delete button
    envRAD$btnUnDel <- gbutton("Un-Delete", handler=VarUnDel, cont=frmApply)          #Un-Delete button
    enabled(envRAD$btnUnDel) <- FALSE                                                 #Disable the Un-Delete button since initially nothing has yet to be deleted 
    fFunc <- gframe("Function Information", cont=grpUpper, expand=TRUE, hor=FALSE)    #Container frame for the Function Info
    envRAD$eFuncType <- glabel("", cont=fFunc)                                        #Label for the Function Type
    envRAD$tFuncDesc <- glabel("", cont=fFunc, wrap=TRUE, editable=FALSE)             #Label for the Function Description
    addSpring(fFunc)                                                                  #Push the next label to the bottom
    envRAD$tInputs <- glabel("", cont=fFunc, wrap=TRUE, editable=FALSE)               #Label for the Input info
    envRAD$tOutputs <- glabel("", cont=fFunc, wrap=TRUE, editable=FALSE)              #Label for the Output info
    envRAD$tFirst <- glabel("", cont=fFunc, wrap=TRUE, editable=FALSE)                #Label for info about the selected out -- currently un-used
    visible(layUpper) <- TRUE                                                         #Now that it's been created, make the upper layout visible
    
    #Middle Panel -- Variables Table
    fVars <- gframe(paste0("Variables for ", basename(envRAD$g.file)), cont=gRAD, expand=TRUE)  #Frame for Variables Table and Up/Down controls
    gVars <- ggroup(cont=fVars, expand=TRUE, spacing=SPACING)                         #Group container for Variables Table
    envRAD$g.tRAD <- gtable(dfVar, cont=gVars, expand=TRUE, multiple=TRUE, handler=VarChanged)  #Table control for Variables
    addHandlerChanged(envRAD$g.tRAD, VarChanged)                                      #Add Handler for when Variables are changed
    addHandlerClicked(envRAD$g.tRAD, VarChanged)                                      #Add Handler for when Variables are clicked
    envRAD$gMove <- glayout(cont=gVars, horizontal=FALSE)                             #Container Layout for Up/Down buttons
    enabled(envRAD$gMove) <- FALSE                                                    #Disable the Up/Down buttons since, initially, no variable is selected (nothing to move)
    envRAD$gMove[1, 1, expand=TRUE] <- gbutton("Up", cont=envRAD$gMove, handler=VarMove)  #Add the Up button
    envRAD$gMove[2, 1, anchor = c(1, 1), expand=TRUE] <- gbutton("Dn", cont=envRAD$gMove, handler=VarMove)  #Add the Down button
    visible(gRAD) <- TRUE                                                             #Make the table visible      
    
    #Bottom Panel -- Metric Expression
    envRAD$gME <- gexpandgroup("Metric Specification", 
                               cont=gRAD, expand=FALSE, spacing=SPACING)              #Container that expands/contracts for Metric Expression
    fOpts <- gframe("Options", cont=envRAD$gME, horizontal=FALSE, spacing=SPACING)                     #Frame container for Options
    gOpt1 <- ggroup(cont=fOpts, expand=FALSE, spacing=SPACING)                              #Group container for Checkboxes (gOpt1A) and Output Variables (gOpt1B)
    gOpt1A <- ggroup(cont=gOpt1, expand=TRUE, horizontal=FALSE, spacing=SPACING)            #Group container for Checkboxes (Options on the left)
    envRAD$chkNul <- gcheckbox("Nulls Allowed", cont=gOpt1A, checked=TRUE, handler=UpdateExpression(NA))     #Checkbox for Allowing Nulls
    envRAD$chkCIC <- gcheckbox("Check Input Count", cont=gOpt1A, checked=TRUE, handler=UpdateExpression(NA)) #Checkbox for Check INput Count
    gOpt1B <- ggroup(cont=gOpt1, horizontal=FALSE, expand=FALSE, spacing=SPACING)           #Group container for Output Variable (Option on the right)
    addSpring(gOpt1B)
    lOut  <- glabel("Output Variable", cont=gOpt1B, anchor=c(-1,0))                   #Label for Output Variable
    envRAD$lstOut <- gdroplist(c("Temp1", "Temp2"), editable=FALSE, cont=gOpt1B, handler=UpdateExpression(NA))  #Dropdown container for Outputs (list to be udpated as Variables are processed and changed)  
    gOpt2 <- ggroup(cont=fOpts, horizontal=FALSE, spacing=SPACING)                                     #Group container for Working Directory (Option on the bottom)
    envRAD$chkSort <- gcheckbox("Enable Sort By", cont=gOpt2, checked=TRUE,  
                                handler=function(h, ...) {                            #Handler for SortBy Checkbox
                                  enabled(envRAD$edtSrt) <- svalue(envRAD$chkSort)    #  Enable SortBy edit box if checked, disable otherwise
                                  envRAD$saved <- FALSE                               #  Remember this change
                                  UpdateExpression(NA)                                #  Update the expression
                                })                                                    #Checkbox for SortBy
    envRAD$edtSrt <- gedit(SORT_BY_DEFAULT, editable=TRUE, expand=TRUE, cont=gOpt2, 
                           handler=function(h, ...) {UpdateExpression(NA)})   #Edit box for SortBy value
    envRAD$chkDir <- gcheckbox("Specify Working Directory", checked=(!is.na(envRAD$workDr)), cont=gOpt2, expand=FALSE, 
                               handler=function(h, ...) {                             #Handler for Working Directory Checkbox
                                 enabled(envRAD$edtDir) <- svalue(envRAD$chkDir)      #  Enable directory edit box if checked, disable otherwise
                                 envRAD$saved <- FALSE                                #  Remember this change
                                 UpdateExpression(NA)                                 #  Update the expression
                               })                                                     #Checkbox for setting the Working Directory and it's handler
    envRAD$edtDir <- gedit("", editable=TRUE, expand=TRUE, cont=gOpt2, 
                           handler=function(h, ...) {UpdateExpression(NA)})                #Edit box for Working Directory
    addHandlerKeystroke(envRAD$edtDir, handler=UpdateExpression(NA))                          #Add handler for editing the working directory
    enabled(envRAD$edtDir) <- (!is.na(envRAD$workDr))                                 #Disable edit box for working Directory, unless there was a working directory in the metric expression since it's checkbox is initially disabled
    fME <- gframe("Metric Expression", cont=envRAD$gME, expand=TRUE, hor=FALSE, spacing=SPACING)       #Container Frame for Metric Expression
    envRAD$txtMetricExp <- gtext("", container=fME, expand=TRUE)                      #Text box control for Metric Expression
    gbutton("Copy to Clipboard", container=fME, handler = function(h, ...) {          #Copy ME to Clipboard Handler
      writeClipboard(svalue(envRAD$txtMetricExp))                                     #  If changes have been made but the script has not been saved, warn the user
      if(!envRAD$saved) galert("Be sure to save your script!", title="Unsaved Changes", delay=3, widget=fME, icon="info")
      })                                                                              #Button to copy Metric Expression to Clipboard, and it's handler
  }  #End-CreateDialog
    
  ### Function to Open a script and parse it to find the potential variables
  ScriptOpen <- function(...) {
    if(DEBUG) print("Function enter: ScriptOpen")
    
    # Get R Script
    if(FILE_CHOOSE) (g.file <- file.choose())                 #For normal use, let's user pick the R Script to open
    else (g.file <- "E:\\Demo\\ARIMA\\ARIMA_Script_new2.R")    #For testing only, opens this specific file without prompting the user
    
    if(exists("g.file")) {
      envRAD$g.file <- g.file                                 #Persist the file name      
      sErrMsg <- tryCatch({  
        #envRAD$g.script <- readLines(fCon, warn=TRUE) #Persist the script for use when saving
        envRAD$g.script <- scan(envRAD$g.file, what=character(), sep="\n", quiet=!DEBUG)  #Persist the script for use when saving
        #scriptParsed <- parser(envRAD$g.file)                 #Parse the R script
        scriptParsed <- parser(text=envRAD$g.script)           #Parse the R script
        scriptErrMsg<-""                                      #If we made it here, then there's no errors to report  
      }, error = function(err) {  
        return(geterrmessage())                               #Report error that was caught
      })
      if(nchar(sErrMsg)>0) {
        gmessage(paste0("The script was not able to be processed due to the following error: ","\n\n", sErrMsg), 
                 title="Error with Script", icon="error")
      } else {                                                #Process the R Script
              
        envRAD$allowUpdates <- FALSE                          #Don't allow updates while the script is being processed        
        envRAD$paramUsed[] <- FALSE                           #Reset any previously used parameters to Un-used
        
        # Tokenize Script
        tokens <- attr(scriptParsed, "data")                  #The script tokens are in this data.frame
        if(DEBUG) TOKENS <<- tokens                           #For debugging, make a global copy of the TOKENS
        #Set columns 
        vars <- character(0)
        dir <- character(0)
        dtype <- character(0)
        ptype <- character(0)
        param <- character(0)
        pVals <- character(0)
        
        #Handle any existing comment blocks
        mstrVarStart1 <- as.numeric(tokens[grep("[[:space:]]*#[[:space:]]*MSTR_VAR_START[[:space:]]*",tokens$text), 1])   #Improvement from Li Zhang -- thank you Li!
        mstrVarEnd1  <- as.numeric(tokens[grep("[[:space:]]*#[[:space:]]*MSTR_VAR_END[[:space:]]*",tokens$text), 1])      #Improvement from Li Zhang -- thank you Li!
        mstrVarStart2 <- as.numeric(tokens[grep("[[:space:]]*#[[:space:]]*MICROSTRATEGY_BEGIN[[:space:]]*",tokens$text), 1])   #Improvement from Li Zhang -- thank you Li!
        mstrVarEnd2  <- as.numeric(tokens[grep("[[:space:]]*#[[:space:]]*MICROSTRATEGY_END[[:space:]]*",tokens$text), 1])      #Improvement from Li Zhang -- thank you Li!
        mstrVarStart <- c(mstrVarStart1, mstrVarStart2)
        mstrVarEnd <- c(mstrVarEnd1, mstrVarEnd2)
        #        mstrWorkDir <- as.numeric(tokens[grep(WORKING_DIR_COMMENT, tokens$text), 1])
        rptCount <- 0    #Set default for Repeat Count
        
        if(length(mstrVarStart)!=length(mstrVarEnd)) {
          gmessage(paste0("Number of #MSTR_VAR_START comments (", length(mstrVarStart), 
                          ") does not equal the number of #MSTR_VAR_END comments (", length(mstrVarEnd), 
                          "). Please fix and try again."), 
                   title="MicroStrategy Variables Comment Block Problems", 
                   icon="error")
        }
        else if(length(mstrVarStart)>0) {  
          #Handle existing comments
          
          mstrVars <- numeric(0)                           #Gather all the variables inside the MicroStrategy Comment Block(s)
          for(i in 1:length(mstrVarStart)) mstrVars <- append(mstrVars, seq(mstrVarStart[i]+1, mstrVarEnd[i]-1))
          mstrVars <- mstrVars[(tokens[mstrVars, "token.desc"]=="COMMENT")
                               &(tokens[mstrVars, "col1"]=="0")
                               &tokens[mstrVars, "text"]!="#"]                      #only look at the non-blank line comments
          varTextAll <- tokens[mstrVars, "text"]
          varText <- sapply(varTextAll, function(x) (unlist(strsplit(x,TAG_METRIC_EXP)))[1])
          metExp <- sapply(varTextAll, function(x) (unlist(strsplit(x,TAG_METRIC_EXP)))[2])
          metExp <- metExp[!is.na(metExp)]
          if(length(metExp)>0) {    #See if we have a metric expression, and if so, get the Working Directory and SortBy values
            envRAD$workDr <- unlist(strsplit(unlist(strsplit(metExp[1], "[_WorkingDir]=\"", fixed=TRUE))[2],"\""))[1]  #Check the first metric expression for the Working Directory, if any
            envRAD$sortBy <- unlist(strsplit(unlist(strsplit(metExp[1], "SortBy=(", fixed=TRUE))[2],")"))[1]  #Check the first metric expression for the SortBy, if any
          } else {
            envRAD$workDr <- NA
            envRAD$sortBy <- NA
          }
          

          #Gather all the comment lines to remove and remove them
          for(i in 1:length(mstrVarEnd)) {                                                       #For each header block end line
            while(envRAD$g.script[mstrVarEnd[i]+1]=="") {                                        #Look to see if the next line is an empty line
              mstrVarEnd[i] <- mstrVarEnd[i]+1                                                   #If so, include the empty line as part of the header block (so it can be removed)
            }
          }
          
          varLines <- numeric(0)                                                                #Start with no lines to remove
          for(i in 1:length(mstrVarStart))                                                      #For each header block
            varLines <- append(varLines, seq(mstrVarStart[i], mstrVarEnd[i]))                   #  Collect the lines to remove
          envRAD$g.script <- envRAD$g.script[-varLines]                                         #Remove the comment lines
          
          #Gather the nature of the variables
          bDis <- sapply(varText, function(x) (length(grep("-d", x, ignore.case=TRUE))>0))      #Flag Disableds
          bOut <- sapply(varText, function(x) (length(grep("-o", x, ignore.case=TRUE))>0))      #Flag Outputs
          bPar <- sapply(varText, function(x) (length(grep("-p", x, ignore.case=TRUE))>0))      #Flag Parameters
          bNum <- sapply(varText, function(x) (length(grep("-n", x, ignore.case=TRUE))>0))      #Flag Numerics
          bStr <- sapply(varText, function(x) (length(grep("-st", x, ignore.case=TRUE))>0))     #Flag Strings
          bSca <- sapply(varText, function(x) (length(grep("-s[^tr]|-s$", x, ignore.case=TRUE))>0)) #Flag Scalars - Improvement from Li Zhang -- thank you Li!

          #Handle Repeat Count
          bIn <- !bOut & !bPar                                                                  #bIn array wiil have TRUE for any inputs
          inputNames <- names(bIn[bIn==TRUE])                                                   #Get the 'names' of the inputs
          firstRpt <- grep(" -r", inputNames, ignore.case=TRUE)[1]                              #Find the first repeat, if any
          if(!is.na(firstRpt)) rptCount <- length(inputNames)-firstRpt+1                        #Set repeat count
          
          #Set the spec of the variables
          vars <-sapply(varText, function(x)(unlist(strsplit(sub("[[:space:]]*#[[:space:]]*","",x),"[[:space:]]+")))[1])  #Get variable names -- Improvement from Li Zhang -- thank you Li!                      
          dir <- ifelse(bOut, "Output", ifelse(bPar, "Parameter", ifelse(bDis, "Disabled", "Input")))                     #Set direction
          dtype <- ifelse(bPar, "", ifelse(bStr, "String", ifelse(bNum, "Numeric", "Default")))                           #Set data type
          ptype <- ifelse(bPar, "", ifelse(bSca, "Scalar", "Vector"))                           #Set parameter type
          param <- sapply(varText, function(x) (GetParam(x)))                                   #Get Parameters
          pVals <- c(rep("", length(param)))
          
        } #Done with existing variables!
                
        
        #Next, Handle new variables from the R Script
        symbols <- as.numeric(rownames(tokens[(tokens$token.desc=="SYMBOL"), ]))                #Get the variables from the script
        excluded <- c("microstrategy.executionFlag", "workingDir", ".GlobalEnv", "MSTR_XEQ")    #Symbols to exclude
        symbols <- symbols[!(tokens[symbols, "text"] %in% excluded)]                            #Exclude certain variables that used only by MicroStrategy, never as inputs, outputs or parameters
        symbols <- symbols[(tokens[abs(symbols-2), "token.desc"]!="'$'")]                       #Only include variable that are NOT members (no '$' to the left of the Symbol)
        vars <- append(vars, tokens[symbols, "text"])                                           #Add the variable names
        bOutputVar <- ((tokens[symbols+1, "token.desc"]=="LEFT_ASSIGN") 
                        | (tokens[symbols+1, "token.desc"]=="EQ_ASSIGN")  
                        | (tokens[abs(symbols-2), "token.desc"]=="RIGHT_ASSIGN"))               #Flag varaiables as outputs
        dir <- append(dir, ifelse(bOutputVar, "Output", "Input"))                               #Add the direction for each variable
        dtype <- append(dtype, rep("Numeric", length(symbols)))                                 #Add the data type for each variable
        ptype <- append(ptype, rep("Vector", length(symbols)))                                  #Add the parameter type for each variable
        param <- append(param, rep("", length(symbols)))                                        #Add the parameter for each variable
        pvals <- append(pVals, rep("", length(symbols)))

        badNames <- grep("$", vars, fixed=TRUE)
        if(length(badNames)>0) {
          sNames <- paste(unlist(vars[badNames]), collapse=", ")
          gmessage(paste0("\"$\" detected in ", length(badNames), ifelse(length(badNames)>1," variables:\n", " variable:\n"), 
                          sNames, "\n\nVariables that use \"$\" are often not in the global environment and only variables in the global environment can be used as inputs and outputs.  Note that variables can be promoted to global by using the \"<<-\" operator or the \"assign\" function."),
                   title="Non-Global Variable Warning", icon="warning")
        }
        
        #Now that we've collected all the variables (old and new), put them into a data frame
        envRAD$dfVar <- data.frame(vars, dir, dtype, ptype, param, pVals, stringsAsFactors=FALSE)      #Create variable data frame
        envRAD$dfVar <- envRAD$dfVar[!duplicated(envRAD$dfVar[ , 1]), ]                         #Delete any duplicate variables, keeping the first one
        colnames(envRAD$dfVar) <- c("Name", "Direction", "Data Type", "Parameter Type", "Parameter", "Value")  #Set the column headers for the variables table
        envRAD$deletedVars <- envRAD$dfVar[1, ]                                                 #Create a data frame for holding deleted variables
        envRAD$delVarMin <- nrow(envRAD$deletedVars)                                            #Remember row count at this time: this value means the deleted var list can be considered empty (no variables to un-delete)
        
        #Add the script and variables data frame to the GUI
        if (!envRAD$loaded) {                                                                   #Is this the first time we've opened a script?
          CreateDialog(envRAD$dfVar)                                                            #This is the first time, so create the dialog
        } else {                                                                                #The dialog was previously created, so
          envRAD$g.tRAD[, ] <- envRAD$dfVar                                                     #  update the variables tables
          svalue(envRAD$lstParam, index=TRUE) <- 1                                              #  Reset the selected parameter to the first blank one
          enabled(envRAD$btnApply) <- FALSE                                                     #  Disable the Apply Button (since no variable has been selected yet)
          enabled(envRAD$lstParam) <- FALSE                                                     #  Disable the Parameter Dropdown since initially the Direction is NOT set to Parameter
          enabled(envRAD$btnUnDel) <- FALSE                                                     #  Disable the UnDelete button since nothing has yet to be deleted
          svalue(envRAD$chkSort) <- TRUE                                                        #  Reset the Sort-By to it's default (TRUE)
          svalue(envRAD$chkCIC) <- TRUE                                                         #  Reset the Check Input Count to it's default (TRUE)
          svalue(envRAD$chkNul) <- TRUE                                                         #  Reset the Nulls Allowed to it's default (TRUE)
        }
        
        if(is.na(envRAD$workDr)) {                                                              #Did the script contain an existing Working Directory?
          svalue(envRAD$chkDir) <- FALSE                                                        #  No, so reset the Working Directory to it's default (FALSE)
          enabled(envRAD$edtDir) <- FALSE                                                       #      and Disable the Working Direcoty edit box
          svalue(envRAD$edtDir) <- gsub("/", "\\", dirname(envRAD$g.file), fixed=TRUE)          #      and set the directory to the default of the R Script's directory
        } else {
          svalue(envRAD$chkDir) <- TRUE                                                         #  Yes, so set the Working Directory to TRUE
          enabled(envRAD$edtDir) <- TRUE                                                        #      and Enable the Working Direcoty edit box
          svalue(envRAD$edtDir) <- envRAD$workDr                                                #      and set the directory 
        }
        if(is.na(envRAD$sortBy)) {                                                              #Did the script contain an existing SortBy value?
          svalue(envRAD$edtSrt) <- SORT_BY_DEFAULT                                              #  No, so use the default
        } else {
          svalue(envRAD$edtSrt) <- envRAD$sortBy                                                #  Yes, so set it
        }
        svalue(envRAD$RptCt) <- rptCount                                                        #Set the repeat count
        svalue(envRAD$edtScr) <- envRAD$g.file                                                  #Display the script file location
        envRAD$allowUpdates <- TRUE                                                             #Allow updates now that the script has been processed
        svalue(envRAD$g.tRAD, index=TRUE) <- 1                                                  #Select the first variable
        SetOutputs()                                                                            #Set the outputs for the first time
        UpdateExpression(NA)                                                                    #Update the metric expression
        envRAD$saved <- TRUE                                                                    #Set the saved flag, we haven't modified any results (yet)
        envRAD$loaded <- TRUE                                                                   #Set the loaded flag, we have loaded an R Script 
        
      } #End of if-else -- no error caught processing Script
      
    } #End of if-else -- file was opened
    
  } #End-ScriptOpen 
  
  if(DEBUG) print("RAD: Handelers Created")
  
  #' @title R Analytic Deployer
  #'
  #' @description
  #' \code{MicroStrategyRAD} launches the MicroStrategy R Analytic Deployer 
  #' to prepare an R Script for execution by MicroStrategy.
  #'
  #' @details
  #' This function launches a user interface which allows the user to 
  #' open an R Script, capture its "signature" (the nature of inputs 
  #' and outputs along with other information about how the R analytic 
  #' should be executed).
  #' 
  #' @param envRAD R environment that contains all information related to the deployment of an R Analytic
  #' @keywords microstrategy
  #' @export
  #' @examples
  #' MicroStrategyRAD(envRAD <- new.env())
  
  MicroStrategyRAD <- function(envRAD) {
  #  if(DEBUG) print("Function enter: MicroStrategyRAD")
    
   ### Create R Analytic Deployer Dialog  
    evalq({
      #Member variables  
      dirList <- c("Input", "Output", "Parameter", "Disabled")
      dtypeList <- c("", "Numeric", "String", "Default")
      ptypeList <- c("", "Vector", "Scalar")
      paramList <- c("", paste("BooleanParam", seq(1, 9), sep = ""), 
                     paste("NumericParam", seq(1, 9), sep = ""), paste("StringParam", seq(1, 9), sep = ""))
      paramUsed <- rep(FALSE, length(paramList))
      names(paramUsed) <- paramList
      funcName <- c("RScriptSimple", "RScriptRelative", "RScriptAgg", 
                    "RScriptRelativeS", "RScriptAggS", "<--INVALID-->", 
                    "<--INVALID-->", "<--INVALID-->", "<--INVALID-->")    
      funcNameNEW <- c("RScriptSimple", "RScriptU", "RScriptAggU", 
                       "RScript", "RScriptAgg", "<--INVALID-->", 
                       "<--INVALID-->", "<--INVALID-->", "<--INVALID-->")    
      funcList <- c("Simple", "Relative (Unsorted)", "Aggregation (Unsorted)", 
                    "Relative", "Aggregation", "<--INVALID-->", 
                    "<--INCOMPLETE-->", "<--SORT BY ERROR-->", "<--OUTPUTS ERROR-->")
      funcDesc <- c("Scalar inputs & output (row at a time)", 
                    "Vector inputs & output (table at a time)", 
                    "Vector inputs & scalar output (grouping)",
                    "Vector inputs & output (table at a time)", 
                    "Vector inputs & scalar output (grouping)", 
                    "Scalar inputs & Vector output are not allowed", 
                    "Need at least one input & one output",
                    "Sort By requires that the first input is a vector",
                    "Outputs must be either all scalar or all vector")
      envRAD$workDr <- NA
      envRAD$sortBy <- NA
      saved <- TRUE
      loaded <- FALSE
      allowUpdates <- FALSE
      currentItem <- -1
      if(DEBUG) print("RAD: Member variables Created")
      
      winRAD <- gwindow("MicroStrategy R Analytic Deployer", spacing=SPACING)       #Container for R Analytic Deployer Window
      size(winRAD) <- c(700, 500)
      gtoolbar(list(open=gaction("Open", icon="open", , handler=function(...) {
        if(!envRAD$saved) {
          gconfirm("Changes have not been saved, opening a new script will causes these changes to be lost. Do you want to save changes before opening a new script?", 
                   title="Save Changes Before Opening Script?", icon="warning", handler=ScriptSave2)
        }
        ScriptOpen()
      }), 
                    save=gaction("Save", icon="save", handler=ScriptSave2), 
                    quit=gaction("Quit", icon="quit", handler=function(...) {
                      if(!envRAD$saved) {
                        gconfirm("Changes have not been saved, do you still want to quit?", 
                                 title="Exit without Saving?", icon="warning", handler=dispose(winRAD))
                      } else dispose(winRAD)
                    }), 
                    about=gaction("About", icon="about", 
                                  handler=function(...) {
                                    gmessage(title="MicroStrategy R Analytic Deployer", 
                                             paste("This utility captures the signature of your R analytic into a header block that's used by MicroStrategy when executing the R Script.", 
                                                   "", "Version 0.9", sep = "\n"))
                                  })
      ), cont = winRAD)                                                              #Define Toolbar
    }, envRAD)
    
    if(DEBUG) print("RAD: Environment and Dialog creation complete")
  
  }  #End-MicroStrategyRAD
  
### Main 
#if(DEBUG) print("-> Launching MicroStrategyRAD")
#envRAD <- new.env()   #Environment for R Analytic Deployer
#MicroStrategyRAD(envRAD)     #Launch the R Analytic Deployer	
#if(DEBUG) print(" <- RAD is finished")
 