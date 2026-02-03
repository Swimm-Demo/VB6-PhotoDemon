---
title: User Action Processing Flow
---
This document outlines the main flow for handling user-initiated actions in the photo editor. When a user selects a command or uses a tool, the system prepares for processing, routes the action to the correct handler, and may prompt for additional input if needed. The requested operation is executed, undo/redo state and UI are updated, and the user receives feedback or results. The flow supports both interactive and automated scenarios, including macro recording and batch processing.

# Processing and Pre-Action State Management

```mermaid
%%{init: {"flowchart": {"defaultRenderer": "elk"}} }%%
flowchart TD
  node1["Start processing user action"]
  click node1 openCode "Modules/Processor.bas:89:142"
  node1 --> node2{"Is this a repeat of last action?"}
  click node2 openCode "Modules/Processor.bas:129:139"
  node2 -->|"Yes"| node3["Restore last action parameters"]
  click node3 openCode "Modules/Processor.bas:130:138"
  node2 -->|"No"| node4["Prepare for processing (UI, parameters, pre-processing)"]
  click node4 openCode "Modules/Processor.bas:142:151"
  node3 --> node4
  node4 --> node5{"Does action require rasterization?"}
  click node5 openCode "Modules/Processor.bas:153:158"
  node5 -->|"User cancels"| node15["Cancel and restore UI"]
  click node15 openCode "Modules/Processor.bas:159:161"
  node5 -->|"No or user agrees"| node6{"Does action require selection removal?"}
  node6 -->|"Yes"| node7["Remove selection"]
  click node7 openCode "Modules/Processor.bas:172:172"
  node6 -->|"No"| node8["Notify macro recorder"]
  click node8 openCode "Modules/Processor.bas:176:176"
  node7 --> node8
  node8 --> node9["File Menu Command Dispatch and Dialog Handling"]
  
  node9 -->|"Found"| node10{"Special case: exit program?"}
  click node10 openCode "Modules/Processor.bas:224:236"
  node10 -->|"Yes"| node11["Exit application"]
  click node11 openCode "Modules/Processor.bas:226:233"
  node10 -->|"No"| node12["Undo/Redo Finalization and Cancellation Handling"]
  
  node9 -->|"Not found"| node13["Edit Menu Command Dispatch and Clipboard Handling"]
  
  node13 -->|"Found"| node12
  node13 -->|"Not found"| node14["Image Menu Command Dispatch and Dialog Routing"]
  
  node14 -->|"Found"| node12
  node14 -->|"Not found"| node16["Layer Command Dispatch and Execution"]
  
  node16 -->|"Found"| node12
  node16 -->|"Not found"| node17["Try Select Menu handler"]
  click node17 openCode "Modules/Processor.bas:250:251"
  node17 -->|"Found"| node12
  node17 -->|"Not found"| node18["Adjustment Command Dispatch and Execution"]
  
  node18 -->|"Found"| node12
  node18 -->|"Not found"| node19["Effect Dispatch and Dialog Routing"]
  
  node19 -->|"Found"| node12
  node19 -->|"Not found"| node20["Try Tools Menu handler"]
  click node20 openCode "Modules/Processor.bas:259:320"
  node20 -->|"Found"| node12
  node20 -->|"Not found"| node21["Show error: unknown action"]
  click node21 openCode "Modules/Processor.bas:323:327"
  node12 --> node22["Finish processing"]
  click node22 openCode "Modules/Processor.bas:391:398"
  node11 --> node22
classDef HeadingStyle fill:#777777,stroke:#333,stroke-width:2px;
click node9 goToHeading "File Menu Command Dispatch and Dialog Handling"
node9:::HeadingStyle
click node13 goToHeading "Edit Menu Command Dispatch and Clipboard Handling"
node13:::HeadingStyle
click node14 goToHeading "Image Menu Command Dispatch and Dialog Routing"
node14:::HeadingStyle
click node16 goToHeading "Layer Command Dispatch and Execution"
node16:::HeadingStyle
click node18 goToHeading "Adjustment Command Dispatch and Execution"
node18:::HeadingStyle
click node19 goToHeading "Effect Dispatch and Dialog Routing"
node19:::HeadingStyle
click node12 goToHeading "Undo/Redo Finalization and Cancellation Handling"
node12:::HeadingStyle
```

<SwmSnippet path="/Modules/Processor.bas" line="89">

---

In `Process`, we start by prepping error handling and logging, then immediately call SetProcessorUI_Busy to lock the UI and prevent user interaction during processing.

```visual basic
Public Sub Process(ByVal processID As String, Optional raiseDialog As Boolean = False, Optional processParameters As String = vbNullString, Optional createUndo As PD_UndoType = UNDO_Nothing, Optional relevantTool As Long = -1, Optional recordAction As Boolean = True)

    'Main error handler for the software processor is initialized by this line
    On Error GoTo MainErrHandler
    
    'Every time this sub is entered, increment the process counter.  You can check for this value being > 1 to see if we are in
    ' the midst of a nested processor request.
    m_NestedProcessingCount = m_NestedProcessingCount + 1
    
    'PD provides several failsafes to avoid unwanted user interaction during processing.  One of these failsafes involves forcibly
    ' removing keyboard focus from our thread.  To ensure that we can properly restore focus when we exit, we cache the currently
    ' focused object prior to disabling it.  (Note that this only triggers on top-level Process calls; nested calls will just
    ' grab the cleared value of "0", which defeats the whole point.)
    Dim procStartTime As Currency
    If (Not raiseDialog) Then
        VBHacks.GetHighResTime procStartTime
        m_FocusHWnd = g_WindowManager.GetFocusAPI
    End If
    
    'Debug mode tracks process calls (as it's a *huge* help when trying to track down unpredictable errors)
    If raiseDialog Then
        PDDebug.LogAction "Show """ & processID & """ dialog", PDM_Processor
    Else
        PDDebug.LogAction """" & processID & """: " & Replace$(processParameters, vbCrLf, vbNullString), PDM_Processor
    End If
    
    'Store the passed parameters inside a local PD_ProcessCall object; some external functions prefer to
    ' receive proc info like this, instead of as separate params.
    Dim thisProcData As PD_ProcessCall
    With thisProcData
        .pcID = processID
        .pcParameters = processParameters
        .pcRaiseDialog = raiseDialog
        .pcRecorded = recordAction
        .pcTool = relevantTool
        .pcUndoType = createUndo
    End With
    
    'If we are simply repeating the last command, replace all the method parameters (which will be blank) with data
    ' from the LastEffectsCall object; this simple approach lets us repeat the last action effortlessly!
    If Strings.StringsEqual(processID, "repeat last action", True) Then
        thisProcData = m_LastProcess
        With m_LastProcess
            processID = .pcID
            raiseDialog = .pcRaiseDialog
            processParameters = .pcParameters
            createUndo = .pcUndoType
            relevantTool = .pcTool
            recordAction = .pcRecorded
        End With
    End If
    
    'Before proceeding, deactivate any interactive UI elements
    SetProcessorUI_Busy processID, raiseDialog, processParameters, createUndo, relevantTool, recordAction
    
```

---

</SwmSnippet>

<SwmSnippet path="/Modules/Processor.bas" line="1256">

---

`SetProcessorUI_Busy` sets up the UI to show it's busy, including the cursor and internal flags, but only if we're actually processing and not just showing a dialog.

```visual basic
Private Sub SetProcessorUI_Busy(ByVal processID As String, Optional raiseDialog As Boolean = False, Optional processParameters As String = vbNullString, Optional createUndo As PD_UndoType = UNDO_Nothing, Optional relevantTool As Long = -1, Optional recordAction As Boolean = True)
    
    'The generic MarkProgramBusyState() function will handle most of this for us, but we first need to figure out if it's appropriate
    ' to do things like display an hourglass cursor.
    Dim useBusyCursor As Boolean: useBusyCursor = False
    
    'If we are modifying the image in some way, and the action is likely to take awhile, display a busy cursor.
    If (Not raiseDialog) Then
        If (createUndo = UNDO_Everything) Or (createUndo = UNDO_Image) Or (createUndo = UNDO_Image_VectorSafe) Or (createUndo = UNDO_Layer) Then useBusyCursor = True
    End If
    
    'Note that the processor is currently running; some UI tasks use this to suspend painting ops.  (If we're just being used as
    ' a thin wrapper to raise a dialog, we'll skip this step, as it's pointless!)
    If (Not raiseDialog) Then m_Processing = True
    
    Processor.MarkProgramBusyState True, useBusyCursor, False, (Not raiseDialog)
    
End Sub
```

---

</SwmSnippet>

<SwmSnippet path="/Modules/Processor.bas" line="144">

---

Back in `Process`, after locking the UI, we call Processor_BeforeStarting to handle quirks like Crop, making sure undo/redo works correctly by prepping the undo stack if needed.

```visual basic
    'Create a parameter parser to handle the parameter string.  This can parse out individual function parameters as specific
    ' data types as necessary.  (Some pre-processing steps require parameter knowledge.)
    Dim cXMLParams As pdSerialize
    Set cXMLParams = New pdSerialize
    If (LenB(processParameters) <> 0) Then cXMLParams.SetParamString processParameters
    
    'A handful of functions (Crop, most notably) require special handling before proceeding.
    Processor_BeforeStarting processID, raiseDialog, processParameters, createUndo, relevantTool, recordAction
    
```

---

</SwmSnippet>

<SwmSnippet path="/Modules/Processor.bas" line="705">

---

`Processor_BeforeStarting` checks if we're about to Crop and, if so, forces the undo manager to save both image and selection state before the crop happens, but only if we're not raising a dialog and not in batch mode.

```visual basic
Public Sub Processor_BeforeStarting(ByVal processID As String, Optional raiseDialog As Boolean = False, Optional processParameters As String = vbNullString, Optional createUndo As PD_UndoType = UNDO_Nothing, Optional relevantTool As Long = -1, Optional recordAction As Boolean = True)
    
    'We need to deal with a strange occurrence before processing PD's "Crop" command.
    ' This command forcibly clears the active selection upon completion.  This is done as a
    ' convenience because after cropping, the active selection is likely misaligned against the
    ' new image.  Unfortunately, this behavior wreaks havoc on PD's Undo/Redo engine, because the
    ' Undo/Redo engine only saves image state *after* an action has completed.  So the image's
    ' state post-Crop is saved nicely, but pre-Crop it may not be, because the selection got
    ' removed out-of-process.
    
    'We also can't remove the selection prior to cropping, because we obviously need its data
    ' to process the crop!
    
    'Thus the need for this workaround.  Prior to applying a crop, we ask the Undo/Redo engine
    ' to forcibly change its previous Undo record to an UNDO_EVERYTHING entry.  This will back
    ' up both the image and selection state prior to the crop, without doing anything
    ' problematic like adding dummy entries to the Undo/Redo chain.
    
    '(Note that the initial "Crop" process (e.g. the one generated by the main menu) requests
    ' raiseDialog as TRUE, even though no dialog is shown.  It does this to trigger some
    ' diagnostic functions that determine whether a non-destructive crop can be applied;
    ' anyway, because of this, we only need to forcibly modify the previous Undo entry if
    ' raiseDialog is FALSE.)
    If (Strings.StringsEqual("Crop", processID, True) And (Not raiseDialog) And (Macros.GetMacroStatus <> MacroBATCH)) Then
        PDImages.GetActiveImage.UndoManager.ForceLastUndoDataToIncludeEverything
    End If
    
End Sub
```

---

</SwmSnippet>

<SwmSnippet path="/Modules/Processor.bas" line="153">

---

Back in `Process`, after prepping undo/redo, we check if the action needs rasterizing vector layers. If the user cancels, we exit early to avoid destructive changes.

```visual basic
    'Next, we need to check for actions that may require us to rasterize one or more vector layers before proceeding.
    ' The process for checking this is rather involved, so we offload it to a separate function.
    '
    'The important thing to note is that a *FALSE* return requires us to immediately exit the processor, as the user has
    ' chosen to cancel the current action.
    If (Not CheckRasterizeRequirements(processID, raiseDialog, processParameters, createUndo)) Then
```

---

</SwmSnippet>

<SwmSnippet path="/Modules/Processor.bas" line="1107">

---

`CheckRasterizeRequirements` checks if the current action will destroy vector data and, if so, prompts the user to rasterize. It handles exceptions for merges and non-destructive crops, and cancels the action if the user says no.

```visual basic
Private Function CheckRasterizeRequirements(ByVal processID As String, Optional raiseDialog As Boolean = False, Optional processParameters As String = vbNullString, Optional createUndo As PD_UndoType = UNDO_Nothing) As Boolean
    
    'Assume that the user is more likely to proceed than cancel, and we will deal with cancellation states as they arise.
    CheckRasterizeRequirements = True
    
    'Some functions require us to parse parameters for additional details; for example, "merge layers" requires us to
    ' check the involved layers to see if they are vector or text layers.
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    cParams.SetParamString processParameters
    
    'If the current layer is a vector layer, and the requested operation is *not* vector-safe, raise a rasterization warning.
    ' This gives the user a chance to back out before permanently ruining the layer.  (Note that the rasterization dialog
    ' offers a "remember my choice" setting, and if that was previously used, we'll skip the dialog portion entirely.)
    '
    '(Also: if this is a showDialog operation, we skip this step, so the user can play around without being bombarded by
    ' rasterization prompts.)
    
    'Start with obvious "this check is pointless" states, like no images being loaded
    If PDImages.IsImageActive() Then
        
        Dim i As Long
        
        Dim okayToRasterize As VbMsgBoxResult
        okayToRasterize = vbCancel
        
        'First, check for the case of operations that modify an entire image (e.g. "Flatten").  Three criteria must be met:
        ' 1) No dialog is being shown
        ' 2) The current layer must contain one or more vector layers
        ' 3) The Undo type must be UNDO_IMAGE or UNDO_EVERYTHING.  Header-only Undo operations (e.g. "Canvas Size") do not
        '    affect vector layers in a destructive manner.
        Dim rasterizeImagePromptNeeded As Boolean
        rasterizeImagePromptNeeded = (Not raiseDialog)
        rasterizeImagePromptNeeded = rasterizeImagePromptNeeded And (PDImages.GetActiveImage.GetNumOfVectorLayers > 0)
        rasterizeImagePromptNeeded = rasterizeImagePromptNeeded And ((createUndo = UNDO_Image) Or (createUndo = UNDO_Everything))
        
        'If this action requires rasterization, let's now check for a few exceptions.
        ' 1) Layer merge operations require us to make a full Undo/Redo copy of the entire image stack, because layer IDs are directly
        '    affected by the result (e.g. one ID goes missing after the merge).  This means they use undo type "UNDO_IMAGE".  However,
        '    if an image contains vector layers, we only need to display a rasterize prompt if one or more of the *merged layers* are
        '    vector layers.  (Merging two raster layers in an image with other vector layers shouldn't display a prompt.)
        '
        ' 2) "Crop image" doesn't need to rasterize vector layers *IF* the selection is rectangular and un-feathered.
        '    (In this case, the crop can be mimicked by just moving layer offsets.)
        '
        'Handle such exceptions now.
        If rasterizeImagePromptNeeded Then
            
            'For each case, determine if a vector layer is being merged, and if not, reset rasterizeImagePromptNeeded.
            ' (These checks must be handled manually, as the layers potentially involved vary by action - e.g. "Merge layer down"
            '  affects different layers than "Merge visible layers".)
            If Strings.StringsEqual(processID, "Merge layer down", True) Then
                If PDImages.GetActiveImage.GetLayerByIndex(cParams.GetLong("layerindex")).IsLayerRaster And PDImages.GetActiveImage.GetLayerByIndex(cParams.GetLong("layerindex") - 1).IsLayerRaster Then
                    rasterizeImagePromptNeeded = False
                End If
                
            ElseIf Strings.StringsEqual(processID, "Merge layer up", True) Then
                If PDImages.GetActiveImage.GetLayerByIndex(cParams.GetLong("layerindex")).IsLayerRaster And PDImages.GetActiveImage.GetLayerByIndex(cParams.GetLong("layerindex") + 1).IsLayerRaster Then
                    rasterizeImagePromptNeeded = False
                End If
            
            ElseIf Strings.StringsEqual(processID, "Merge visible layers", True) Then
                
                rasterizeImagePromptNeeded = False
                For i = 1 To PDImages.GetActiveImage.GetNumOfLayers - 1
                    
                    'If a vector layer is found, restore rasterizeImagePromptNeeded and exit the loop
                    If PDImages.GetActiveImage.GetLayerByIndex(i).GetLayerVisibility And PDImages.GetActiveImage.GetLayerByIndex(i).IsLayerVector Then
                        rasterizeImagePromptNeeded = True
                        Exit For
                    End If
                
                Next i
            
            ElseIf Strings.StringsEqual(processID, "Crop", True) Then
                If cParams.GetBool("nondestructive", False, True) Then rasterizeImagePromptNeeded = False
            End If
            
        End If
        
        'If we need to do a "special-case whole-image" rasterization, do so now.
        If rasterizeImagePromptNeeded Then
            
            okayToRasterize = Layers.AskIfOkayToRasterizeLayer(PDImages.GetActiveImage.GetActiveLayer.GetLayerType, , True)
            If (okayToRasterize = vbYes) Then
                
                'When merging layers, only the merged layers need to be rasterized.  (We want to perform as few rasterizations
                ' as possible, so we manually handle each merge case specially.)
                If Strings.StringsEqual(processID, "Merge layer down", True) Then
                    If PDImages.GetActiveImage.GetLayerByIndex(cParams.GetLong("layerindex")).IsLayerVector Then Layers.RasterizeLayer cParams.GetLong("layerindex")
                    If PDImages.GetActiveImage.GetLayerByIndex(cParams.GetLong("layerindex") - 1).IsLayerVector Then Layers.RasterizeLayer cParams.GetLong("layerindex") - 1
                    
                ElseIf Strings.StringsEqual(processID, "Merge layer up", True) Then
                    If PDImages.GetActiveImage.GetLayerByIndex(cParams.GetLong("layerindex")).IsLayerVector Then Layers.RasterizeLayer cParams.GetLong("layerindex")
                    If PDImages.GetActiveImage.GetLayerByIndex(cParams.GetLong("layerindex") + 1).IsLayerVector Then Layers.RasterizeLayer cParams.GetLong("layerindex") + 1
                    
                ElseIf Strings.StringsEqual(processID, "Merge visible layers", True) Then
                    For i = 1 To PDImages.GetActiveImage.GetNumOfLayers - 1
                        If PDImages.GetActiveImage.GetLayerByIndex(i).GetLayerVisibility And PDImages.GetActiveImage.GetLayerByIndex(i).IsLayerVector Then
                            Layers.RasterizeLayer i
                        End If
                    Next i
                        
                'For any other case, rasterize all vector layers
                Else
                    Layers.RasterizeLayer -1
                End If
                
            'If the user doesn't want rasterization, bail immediately.
            Else
                CheckRasterizeRequirements = False
            End If
            
        End If
        
        'At this point, we have dealt with "full-image" modifications - like "Flatten" or "Merge layers" - that may require rasterization.
        
        'Next, we want to deal with operations that modify just *one* layer.  (These are much easier to handle.)
        If CheckRasterizeRequirements And (Not rasterizeImagePromptNeeded) Then
            
            rasterizeImagePromptNeeded = (Not raiseDialog)
            rasterizeImagePromptNeeded = rasterizeImagePromptNeeded And PDImages.GetActiveImage.GetActiveLayer.IsLayerVector
            rasterizeImagePromptNeeded = rasterizeImagePromptNeeded And (createUndo = UNDO_Layer)
            
            'As before, display a "do you want to rasterize?" prompt as necessary
            If rasterizeImagePromptNeeded Then
                
                okayToRasterize = Layers.AskIfOkayToRasterizeLayer(PDImages.GetActiveImage.GetActiveLayer.GetLayerType)
                
                'If rasterization is okay, apply it immediately
                If (okayToRasterize = vbYes) Then
                    Layers.RasterizeLayer PDImages.GetActiveImage.GetActiveLayerIndex
                
                'If the user doesn't want rasterization, bail immediately.
                Else
                    CheckRasterizeRequirements = False
                End If
                
            End If
            
        End If
        
    End If
    
End Function
```

---

</SwmSnippet>

<SwmSnippet path="/Modules/Processor.bas" line="159">

---

Back in `Process`, if rasterization was cancelled, we call SetProcessorUI_Idle to unlock the UI and restore focus, then exit early.

```visual basic
        SetProcessorUI_Idle processID, raiseDialog, processParameters, createUndo, relevantTool, recordAction
        Exit Sub
    End If
    
```

---

</SwmSnippet>

<SwmSnippet path="/Modules/Processor.bas" line="1275">

---

`SetProcessorUI_Idle` unlocks the UI, updates the nested counter, and restores focus if needed.

```visual basic
Private Sub SetProcessorUI_Idle(ByVal processID As String, Optional raiseDialog As Boolean = False, Optional processParameters As String = vbNullString, Optional createUndo As PD_UndoType = UNDO_Nothing, Optional relevantTool As Long = -1, Optional recordAction As Boolean = True)
    
    m_Processing = False
    m_NestedProcessingCount = m_NestedProcessingCount - 1
    Processor.MarkProgramBusyState False, True, False
    
    'Manually handle focus restoration
    If (m_NestedProcessingCount = 0) And (m_FocusHWnd <> 0) Then
        If (Not g_WindowManager Is Nothing) Then g_WindowManager.SetFocusAPI m_FocusHWnd
        m_FocusHWnd = 0
    End If
    
End Sub
```

---

</SwmSnippet>

<SwmSnippet path="/Modules/Processor.bas" line="163">

---

Back in `Process`, after restoring the UI, we check if the current action needs the selection removed in advance (like resizing or rotating), and trigger removal if needed to keep things in sync.

```visual basic
    'If a selection is active, certain functions (primarily transformations) will remove it before proceeding.
    ' This is typically done by functions that resize or reorient the image in a way that makes the selection's
    ' shape irrelevant. Because PD requires the selection mask and image size to remain in sync, errors may occur
    ' if selections persist after a size change - and this is particularly relevant for the Undo/Redo engine,
    ' because it will crash if it attempts to load an Undo file of an image, and the image size is not the same
    ' as the current selection.
    '
    'Anyway, before moving deeper into the processor, check for actions that disallow selections, and prior to
    ' processing them, initiate a Remove Selection request.
    RemoveSelectionAsNecessary processID, raiseDialog, processParameters, createUndo
    
```

---

</SwmSnippet>

<SwmSnippet path="/Modules/Processor.bas" line="1059">

---

`RemoveSelectionAsNecessary` checks if a selection is active and if the action is one that changes image geometry. If so, it calls Processor.Process to remove the selection before proceeding.

```visual basic
Private Function RemoveSelectionAsNecessary(ByVal processID As String, Optional raiseDialog As Boolean = False, Optional processParameters As String = vbNullString, Optional createUndo As PD_UndoType = UNDO_Nothing) As Boolean

    If (Not raiseDialog) And PDImages.IsImageActive() Then
    
        'Only worry about this step if a selection is currently active
        If PDImages.GetActiveImage.IsSelectionActive And (createUndo <> UNDO_Selection) Then
    
            Dim removeSelectionInAdvance As Boolean
            removeSelectionInAdvance = False
            
            'If this action reorients or resizes the image, mark the selection for removal
            removeSelectionInAdvance = removeSelectionInAdvance Or Strings.StringsEqual("Resize image", processID, True)
            removeSelectionInAdvance = removeSelectionInAdvance Or Strings.StringsEqual("Resize", processID, True)
            removeSelectionInAdvance = removeSelectionInAdvance Or Strings.StringsEqual("Content-aware image resize", processID, True)
            removeSelectionInAdvance = removeSelectionInAdvance Or Strings.StringsEqual("Canvas size", processID, True)
            removeSelectionInAdvance = removeSelectionInAdvance Or Strings.StringsEqual("Fit canvas to active layer", processID, True)
            removeSelectionInAdvance = removeSelectionInAdvance Or Strings.StringsEqual("Fit canvas around all layers", processID, True)
            removeSelectionInAdvance = removeSelectionInAdvance Or Strings.StringsEqual("Trim empty borders", processID, True)
            removeSelectionInAdvance = removeSelectionInAdvance Or Strings.StringsEqual("Rotate image 90 clockwise", processID, True)
            removeSelectionInAdvance = removeSelectionInAdvance Or Strings.StringsEqual("Rotate 90 clockwise", processID, True)
            removeSelectionInAdvance = removeSelectionInAdvance Or Strings.StringsEqual("Rotate image 180", processID, True)
            removeSelectionInAdvance = removeSelectionInAdvance Or Strings.StringsEqual("Rotate 180", processID, True)
            removeSelectionInAdvance = removeSelectionInAdvance Or Strings.StringsEqual("Rotate image 90 counter-clockwise", processID, True)
            removeSelectionInAdvance = removeSelectionInAdvance Or Strings.StringsEqual("Rotate 90 counter-clockwise", processID, True)
            removeSelectionInAdvance = removeSelectionInAdvance Or Strings.StringsEqual("Arbitrary image rotation", processID, True)
            removeSelectionInAdvance = removeSelectionInAdvance Or Strings.StringsEqual("Arbitrary rotation", processID, True)
            removeSelectionInAdvance = removeSelectionInAdvance Or Strings.StringsEqual("Flip image vertically", processID, True)
            removeSelectionInAdvance = removeSelectionInAdvance Or Strings.StringsEqual("Flip vertically", processID, True)
            removeSelectionInAdvance = removeSelectionInAdvance Or Strings.StringsEqual("Flip image horizontally", processID, True)
            removeSelectionInAdvance = removeSelectionInAdvance Or Strings.StringsEqual("Flip horizontally", processID, True)
            
            'If selection removal is required, process the removal before proceeding with the
            ' original process request
            If removeSelectionInAdvance Then
                RemoveSelectionAsNecessary = True
                Processor.Process "Remove selection", , , UNDO_Selection
            End If
            
        End If
        
    End If
    
End Function
```

---

</SwmSnippet>

<SwmSnippet path="/Modules/Processor.bas" line="174">

---

Back in `Process`, after prepping selection state, we notify the macro recorder and prep undo/redo. Before dispatching the main action, we check for unsaved canvas changes and create undo entries if needed.

```visual basic
    'If we made it all the way here, notify the macro recorder that something interesting has happened.
    ' (It may choose to store this action for later playback.)
    Macros.NotifyProcessorEvent thisProcData
    
    'If a dialog is being displayed, forcibly disable Undo creation.  (This is really just a failsafe; PD's various dialog functions
    ' are smart about not requesting Undo/Redo events for dialog actions.)
    If raiseDialog Then createUndo = UNDO_Nothing
    
    'If this action requires us to create an Undo entry, do so now.  (We can also use this identifier to initiate a few
    ' other, related actions.)
    If (createUndo <> UNDO_Nothing) Then
        
        'Save this action's information in the m_LastProcess variable (to be used if the user clicks on Edit -> Redo Last Action)
        If Actions.IsActionRepeatable(processID, True) Then m_LastProcess = thisProcData
        
        'If the user wants us to time how long this action takes, mark the current time now
        If g_DisplayTimingReports Then VBHacks.GetHighResTime m_ProcessingTime
        
        'Finally, perform a check for any on-canvas modifications that have not yet had their Undo data saved.
        CheckForCanvasModifications createUndo
        
    End If
    
    Dim procSortStartTime As Currency
    If (Not raiseDialog) Then VBHacks.GetHighResTime procSortStartTime
    
    
```

---

</SwmSnippet>

<SwmSnippet path="/Modules/Processor.bas" line="1006">

---

`CheckForCanvasModifications` checks if the selection has changed since the last undo entry by comparing XML strings. If it has, it creates a new undo entry for the selection modification.

```visual basic
Private Sub CheckForCanvasModifications(ByVal createUndo As PD_UndoType)

    On Error GoTo CheckForCanvasModifyFail

    If PDImages.IsImageActive() Then
    
        If PDImages.GetActiveImage.IsSelectionActive And (createUndo <> UNDO_Selection) And (createUndo <> UNDO_Everything) Then
        
            'Ask the Undo engine to return the last selection param string it has on file
            Dim lastSelParamString As String
            lastSelParamString = PDImages.GetActiveImage.UndoManager.GetLastParamString(UNDO_Selection)
            
            'If such a param string exists, compare it against the current selection param string
            If (LenB(lastSelParamString) <> 0) Then
                
                'If the last selection Undo param string does not match the current selection param string, the user has
                ' modified the selection in some way since the last Undo was created.  Create a new entry now.
                If Strings.StringsNotEqual(lastSelParamString, PDImages.GetActiveImage.MainSelection.GetSelectionAsXML, True) Then
                    
                    'Ensure "modify selection" is available to the translation engine
                    Dim tmpString As String
                    tmpString = g_Language.TranslateMessage("Modify selection")
                    
                    Dim tmpProcData As PD_ProcessCall
                    With tmpProcData
                        .pcID = "Modify selection"
                        .pcParameters = PDImages.GetActiveImage.MainSelection.GetSelectionAsXML()
                        .pcRaiseDialog = False
                        .pcRecorded = True
                        .pcUndoType = UNDO_Selection
                    End With
                    
                    PDImages.GetActiveImage.UndoManager.CreateUndoData tmpProcData
                    
                End If
            
            End If
        
        End If
        
    End If
    
    Exit Sub
    
CheckForCanvasModifyFail:
    PDDebug.LogAction "WARNING!  Processor.CheckForCanvasModifications failed unexpectedly (#" & Err.Number & ", " & Err.Description & ")"
    
End Sub
```

---

</SwmSnippet>

<SwmSnippet path="/Modules/Processor.bas" line="201">

---

Back in `Process`, after prepping undo/redo and macro state, we dispatch the processID to the right handler (like Process_FileMenu) or handle legacy/special cases directly. This centralizes all action routing.

```visual basic
    '******************************************************************************************************************
    '
    'BEGIN PROCESS SORTING
    '
    'The bulk of this routine starts here.  From this point on, the processID string is compared against a hard-coded
    ' list of every possible PhotoDemon action, filter, or other operation.  Depending on the processID, additional
    ' actions will be performed.
    '
    'For ease of reference, the various processIDs are divided into categories of similar functions.  These categories
    ' match the organization of PhotoDemon's menus.  Please note that such organization is simply to improve
    ' readability; there are no functional implications.
    '
    '******************************************************************************************************************
    
    'File menu operations have been successfully migrated to XML strings
    Dim processFound As Boolean, returnDetails As String
    processFound = Process_FileMenu(processID, raiseDialog, processParameters, createUndo, relevantTool, recordAction, returnDetails)
    
```

---

</SwmSnippet>

## File Menu Command Dispatch and Dialog Handling

```mermaid
%%{init: {"flowchart": {"defaultRenderer": "elk"}} }%%
flowchart TD
    node1["User selects a File menu command"]
    click node1 openCode "Modules/Processor.bas:1358:1467"
    node1 --> node2{"Which command?"}
    click node2 openCode "Modules/Processor.bas:1360:1465"
    node2 -->|"New image"| node3{"Show dialog?"}
    click node3 openCode "Modules/Processor.bas:1361:1362"
    node3 -->|"Yes"| node4["Show New Image dialog"]
    click node4 openCode "Modules/Interface.bas:993:1095"
    node3 -->|"No"| node5["Create new image"]
    click node5 openCode "Modules/Processor.bas:1361:1362"
    node2 -->|"Open"| node6["Open image"]
    click node6 openCode "Modules/Processor.bas:1364:1366"
    node2 -->|"Close"| node7["Close image"]
    click node7 openCode "Modules/Processor.bas:1368:1370"
    node2 -->|"Close all"| node8["Close all images"]
    click node8 openCode "Modules/Processor.bas:1372:1374"
    node2 -->|"Save/Save as/Save copy"| node9["Save image"]
    click node9 openCode "Modules/Processor.bas:1376:1386"
    node2 -->|"Revert"| node10{"Is revert enabled?"}
    click node10 openCode "Modules/Processor.bas:1388:1393"
    node10 -->|"Yes"| node11["Revert to last saved state"]
    click node11 openCode "Modules/Processor.bas:1389:1391"
    node10 -->|"No"| node25["Operation completed"]
    node2 -->|"Export"| node13["Export image/layers/animation/color lookup/profile/palette"]
    click node13 openCode "Modules/Processor.bas:1395:1423"
    node2 -->|"Batch wizard"| node14["Show batch wizard dialog"]
    click node14 openCode "Modules/Processor.bas:1425:1427"
    node2 -->|"Print"| node15{"Show dialog?"}
    click node15 openCode "Modules/Processor.bas:1430:1441"
    node15 -->|"Yes"| node16["Show print dialog"]
    click node16 openCode "Modules/Interface.bas:993:1095"
    node15 -->|"No"| node17["Print image"]
    click node17 openCode "Modules/Processor.bas:1435:1436"
    node2 -->|"Exit program"| node18["Signal program exit"]
    click node18 openCode "Modules/Processor.bas:1444:1447"
    node2 -->|"Scanner/camera"| node19["Select scanner/camera or scan image"]
    click node19 openCode "Modules/Processor.bas:1449:1455"
    node2 -->|"Screen capture"| node20{"Show dialog?"}
    click node20 openCode "Modules/Processor.bas:1458:1459"
    node20 -->|"Yes"| node21["Show screen capture dialog"]
    click node21 openCode "Modules/Interface.bas:993:1095"
    node20 -->|"No"| node22["Capture screen"]
    click node22 openCode "Modules/Processor.bas:1459:1459"
    node2 -->|"Internet import"| node23{"Show dialog?"}
    click node23 openCode "Modules/Processor.bas:1461:1462"
    node23 -->|"Yes"| node24["Show internet import dialog"]
    click node24 openCode "Modules/Interface.bas:993:1095"
    node2 --> node25["Operation completed"]
    click node25 openCode "Modules/Processor.bas:1362:1467"

classDef HeadingStyle fill:#777777,stroke:#333,stroke-width:2px;
```

<SwmSnippet path="/Modules/Processor.bas" line="1358">

---

In `Process_FileMenu`, we match the processID to known file commands. For commands needing user input, we show a dialog (handled by ShowPDDialog); otherwise, we run the action directly. The function returns True if it handled the command.

```visual basic
Private Function Process_FileMenu(ByVal processID As String, Optional raiseDialog As Boolean = False, Optional processParameters As String = vbNullString, Optional createUndo As PD_UndoType = UNDO_Nothing, Optional relevantTool As Long = -1, Optional recordAction As Boolean = True, Optional ByRef returnDetails As String = vbNullString) As Boolean

    If Strings.StringsEqual(processID, "New image", True) Then
        If raiseDialog Then ShowPDDialog vbModal, FormNewImage Else FileMenu.CreateNewImage processParameters
        Process_FileMenu = True
        
    ElseIf Strings.StringsEqual(processID, "Open", True) Then
        FileMenu.MenuOpen
        Process_FileMenu = True
    
    ElseIf Strings.StringsEqual(processID, "Close", True) Then
        FileMenu.MenuClose
        Process_FileMenu = True
    
    ElseIf Strings.StringsEqual(processID, "Close all", True) Then
        FileMenu.MenuCloseAll
        Process_FileMenu = True
    
    ElseIf Strings.StringsEqual(processID, "Save", True) Then
        FileMenu.MenuSave PDImages.GetActiveImage()
        Process_FileMenu = True
    
    ElseIf Strings.StringsEqual(processID, "Save as", True) Then
        FileMenu.MenuSaveAs PDImages.GetActiveImage()
        Process_FileMenu = True
        
    ElseIf Strings.StringsEqual(processID, "Save copy", True) Then
        FileMenu.MenuSaveLosslessCopy PDImages.GetActiveImage()
        Process_FileMenu = True
        
    ElseIf Strings.StringsEqual(processID, "Revert", True) Then
        If Menus.IsMenuEnabled("file_revert") Then
            PDImages.GetActiveImage.UndoManager.RevertToLastSavedState
            Interface.NotifyImageChanged PDImages.GetActiveImageID()
        End If
        Process_FileMenu = True
    
    ElseIf Strings.StringsEqual(processID, "Export image", True) Then
        FileMenu.MenuExportImage PDImages.GetActiveImage()
        Process_FileMenu = True
    
    ElseIf Strings.StringsEqual(processID, "Export layers", True) Then
        If Menus.IsMenuEnabled("file_export_layers") Then
            If raiseDialog Then
                ShowPDDialog vbModal, FormExportLayers
            Else
```

---

</SwmSnippet>

<SwmSnippet path="/Modules/Interface.bas" line="993">

---

`ShowPDDialog` loads the dialog, sets ownership (main form or previous dialog), positions it (centered or using saved position), mirrors window icons for consistency, and restores everything after the dialog closes.

```visual basic
Public Sub ShowPDDialog(ByRef dialogModality As FormShowConstants, ByRef dialogForm As Form, Optional ByVal doNotUnload As Boolean = False)

    On Error GoTo ShowPDDialogError
    
    m_ModalDialogActive = True
    
    'Make sure PD's main form is visible
    If (FormMain.WindowState = vbMinimized) Then FormMain.WindowState = vbNormal
    
    'Turn off any async pipe connections or other listeners
    FormMain.ChangeSessionListenerState False
    
    'Reset our "last dialog result" tracker.  (We use "ignore" as the "default" value, as it's a value PD never utilizes internally.)
    m_LastShowDialogResult = vbIgnore
    
    'Start by loading the form and hiding it
    If (dialogForm Is Nothing) Then Load dialogForm
    dialogForm.Visible = False
    
    'Store a reference to this dialog; if subsequent dialogs are loaded, this dialog will be given ownership over them
    If (currentDialogReference Is Nothing) Then
        
        'This is a regular modal dialog, and the main form should be its owner
        isSecondaryDialog = False
        Set currentDialogReference = dialogForm
                
    Else
    
        'We already have a reference to a modal dialog - that means a modal dialog is raising *another* modal dialog.  Give the previous
        ' modal dialog ownership over this new dialog!
        isSecondaryDialog = True
        
    End If
    
    'Retrieve and cache the hWnd; we need access to this even if the form is unloaded, so we can properly deregister it
    ' with the window manager.
    Dim dialogHWnd As Long
    dialogHWnd = dialogForm.hWnd
    
    'If the window has a previous position stored, use that.
    Dim prevPositionStored As Boolean
    If (Not g_WindowManager Is Nothing) Then prevPositionStored = g_WindowManager.IsPreviousPositionStored(dialogForm)
    
    'If a previous position is *not* stored, center it against the main dialog.
    If (Not prevPositionStored) Then
        
        'Get the rect of the main form, which we will use to calculate a center position
        Dim ownerRect As winRect
        GetWindowRect FormMain.hWnd, ownerRect
        
        'Determine the center of that rect
        Dim centerX As Long, centerY As Long
        centerX = ownerRect.x1 + (ownerRect.x2 - ownerRect.x1) \ 2
        centerY = ownerRect.y1 + (ownerRect.y2 - ownerRect.y1) \ 2
        
        'Get the rect of the child dialog
        Dim dialogRect As winRect
        GetWindowRect dialogHWnd, dialogRect
        
        'Determine an upper-left point for the dialog based on its size
        Dim newLeft As Long, newTop As Long
        newLeft = centerX - ((dialogRect.x2 - dialogRect.x1) \ 2)
        newTop = centerY - ((dialogRect.y2 - dialogRect.y1) \ 2)
        
        'If this position results in the dialog sitting off-screen, move it so that its bottom-right corner is always on-screen.
        ' (All PD dialogs have bottom-right OK/Cancel buttons, so that's the most important part of the dialog to show.)
        If newLeft + (dialogRect.x2 - dialogRect.x1) > g_Displays.GetDesktopRight Then newLeft = g_Displays.GetDesktopRight - (dialogRect.x2 - dialogRect.x1)
        If newTop + (dialogRect.y2 - dialogRect.y1) > g_Displays.GetDesktopBottom Then newTop = g_Displays.GetDesktopBottom - (dialogRect.y2 - dialogRect.y1)
        
        'Move the dialog into place, but do not repaint it (that will be handled in a moment by the .Show event)
        MoveWindow dialogHWnd, newLeft, newTop, dialogRect.x2 - dialogRect.x1, dialogRect.y2 - dialogRect.y1, 0
        
    End If
    
    'Mirror the current run-time window icons to the dialog; this allows the icons to appear in places like Alt+Tab
    ' on older OSes, even though a toolbox window has focus.
    Interface.FixPopupWindow dialogHWnd, True
    
    'Use VB to actually display the dialog.  Note that code execution will pause here until the form is closed.
    ' (As usual, disclaimers apply to message-loop functions like DoEvents.)
    dialogForm.Show dialogModality, FormMain
    
    'Now that the dialog has finished, we must replace the windows icons with its original ones -
    ' otherwise, VB will mistakenly unload our custom icons with the window!
    Interface.FixPopupWindow dialogHWnd, False
    
    'Release our reference to this dialog
    If isSecondaryDialog Then
        isSecondaryDialog = False
    Else
        Set currentDialogReference = Nothing
    End If
    
    'If the form has not been unloaded, unload it now
    If (Not (dialogForm Is Nothing)) And (Not doNotUnload) Then
        Unload dialogForm
        Set dialogForm = Nothing
    End If
    
    'Reinstate any async listeners
    FormMain.ChangeSessionListenerState True
    
    m_ModalDialogActive = False
    
    Exit Sub
    
'For reasons I can't yet ascertain, this function will sometimes fail, claiming that a modal window is already active.  If that happens,
' we can just exit.
ShowPDDialogError:

    m_ModalDialogActive = False

End Sub
```

---

</SwmSnippet>

<SwmSnippet path="/Modules/Processor.bas" line="1404">

---

After returning from `ShowPDDialog`, `Process_FileMenu` finishes up by handling any remaining commands, including special cases like exit (using returnDetails), and returns True if it processed something.

```visual basic
                'There is no else; the above dialog handles everything!
            End If
        End If
        Process_FileMenu = True
    
    ElseIf Strings.StringsEqual(processID, "Export animation", True) Then
        Saving.Export_Animation PDImages.GetActiveImage()
        Process_FileMenu = True
    
    ElseIf Strings.StringsEqual(processID, "Export color lookup", True) Then
        Saving.SaveColorLookupToFile PDImages.GetActiveImage()
        Process_FileMenu = True
        
    ElseIf Strings.StringsEqual(processID, "Export color profile", True) Then
        ColorManagement.SaveImageProfileToFile PDImages.GetActiveImage()
        Process_FileMenu = True
    
    ElseIf Strings.StringsEqual(processID, "Export palette", True) Then
        Palettes.ExportCurrentImagePalette PDImages.GetActiveImage()
        Process_FileMenu = True
    
    ElseIf Strings.StringsEqual(processID, "Batch wizard", True) Then
        Interface.ShowPDDialog vbModal, FormBatchWizard
        Process_FileMenu = True
             
    ElseIf Strings.StringsEqual(processID, "Print", True) Then
        If raiseDialog Then
            
            'As a temporary workaround, Vista+ users are routed through the default Windows photo printing
            ' dialog.  XP users get the old PD print dialog.
            If OS.IsVistaOrLater Then
                Printing.PrintViaWindowsPhotoPrinter
            Else
                If (Not FormPrint.Visible) Then Interface.ShowPDDialog vbModal, FormPrint
            End If
            
        End If
        Process_FileMenu = True
            
    ElseIf Strings.StringsEqual(processID, "Exit program", True) Then
        
        'The main process function handles this step; we just need to notify it that an exit has been triggered
        returnDetails = PD_PROCESS_EXIT_NOW
        Process_FileMenu = True
        
    ElseIf Strings.StringsEqual(processID, "Select scanner or camera", True) Then
        Plugin_EZTwain.Twain32SelectScanner
        Process_FileMenu = True
            
    ElseIf Strings.StringsEqual(processID, "Scan image", True) Then
        Plugin_EZTwain.Twain32Scan
        Process_FileMenu = True
            
    ElseIf Strings.StringsEqual(processID, "Screen capture", True) Then
        If raiseDialog Then Interface.ShowPDDialog vbModal, FormScreenCapture Else ScreenCapture.CaptureScreen processParameters
        Process_FileMenu = True
        
    ElseIf Strings.StringsEqual(processID, "Internet import", True) Then
        If raiseDialog Then Interface.ShowPDDialog vbModal, FormInternetImport
        Process_FileMenu = True
        
    End If
    
End Function
```

---

</SwmSnippet>

## Edit Menu Command Routing

```mermaid
%%{init: {"flowchart": {"defaultRenderer": "elk"}} }%%
flowchart TD
    node1{"Was a menu operation matched?"}
    click node1 openCode "Modules/Processor.bas:221:238"
    node1 -->|"Yes"| node2{"Is the operation 'exit program'?"}
    click node2 openCode "Modules/Processor.bas:224:236"
    node2 -->|"Yes"| node3{"Did user confirm exit?"}
    click node3 openCode "Modules/Processor.bas:231:234"
    node3 -->|"Yes"| node4["Exit PhotoDemon immediately and stop further processing"]
    click node4 openCode "Modules/Processor.bas:226:233"
    node3 -->|"No"| node8["Continue normal processing"]
    click node8 openCode "Modules/Processor.bas:234:236"
    node2 -->|"No"| node8
    node1 -->|"No"| node5["Process Edit Menu operations"]
    click node5 openCode "Modules/Processor.bas:241:241"
classDef HeadingStyle fill:#777777,stroke:#333,stroke-width:2px;
```

<SwmSnippet path="/Modules/Processor.bas" line="219">

---

Back in `Process`, after FileMenu, we check if the command was handled. If not, we pass it to Process_EditMenu to see if it's an edit command.

```visual basic
    'The File menu contains some abnormal operations (e.g. "exit program") which require us to deal with their return
    ' codes immediately.
    If processFound Then
        
        'The "exit program" menu item requires us to close PhotoDemon immediately; check the returnDetails string for this case
        If Strings.StringsEqual(returnDetails, PD_PROCESS_EXIT_NOW, True) Then
        
            Unload FormMain
            
            'If the user allows the exit to proceed (e.g. they don't hit "cancel"), we must forcibly exit this sub immediately.
            ' (Otherwise, later operations in this function will attempt to access things like FormMain, which are in the midst
            ' of unloading!)
            If g_ProgramShuttingDown Then
                m_NestedProcessingCount = m_NestedProcessingCount - 1
                Exit Sub
            End If
        
        End If
        
    End If
    
    'Edit menu operations have been successfully migrated to XML strings.  (None of their functions raise special return conditions, FYI.)
    If (Not processFound) Then processFound = Process_EditMenu(processID, raiseDialog, processParameters, createUndo, relevantTool, recordAction, returnDetails)
    
```

---

</SwmSnippet>

## Edit Menu Command Dispatch and Clipboard Handling

```mermaid
%%{init: {"flowchart": {"defaultRenderer": "elk"}} }%%
flowchart TD
    node1["User selects an Edit menu action (processID)"] --> node2{"What type of action?"}
    click node1 openCode "Modules/Processor.bas:1472:1478"
    node2 -->|"Undo/Redo/Undo history"| node3["Restore or move to previous image state"]
    click node2 openCode "Modules/Processor.bas:1478:1513"
    node2 -->|"Fade"| node4["Apply fade to last action"]
    click node4 openCode "Modules/Processor.bas:1514:1517"
    node2 -->|"Clipboard operation (Cut, Copy, Paste, etc.)"| node5["Perform clipboard operation (may show dialog)"]
    click node5 openCode "Modules/Processor.bas:1518:1586"
    node2 -->|"Selection operation (Clear, Fill, Stroke, etc.)"| node6["Perform selection operation (may show dialog)"]
    click node6 openCode "Modules/Processor.bas:1588:1603"
    node3 --> node7{"Was Undo/Redo used?"}
    click node7 openCode "Modules/Processor.bas:1606:1612"
    node7 -->|"Yes"| node8["Sync layer properties to active layer"]
    click node8 openCode "Modules/Processor.bas:1609:1611"
    node7 -->|"No"| node9["End"]
    node4 --> node9
    node5 --> node9
    node6 --> node9
classDef HeadingStyle fill:#777777,stroke:#333,stroke-width:2px;
```

<SwmSnippet path="/Modules/Processor.bas" line="1472">

---

In `Process_EditMenu`, we match the processID to edit commands like undo/redo and clipboard ops. For some, we show dialogs (handled by ShowPDDialog); for others, we run the action and sync the UI and layer settings.

```visual basic
Private Function Process_EditMenu(ByRef processID As String, Optional ByVal raiseDialog As Boolean = False, Optional ByRef processParameters As String = vbNullString, Optional ByRef createUndo As PD_UndoType = UNDO_Nothing, Optional ByRef relevantTool As Long = -1, Optional ByRef recordAction As Boolean = True, Optional ByRef returnDetails As String = vbNullString) As Boolean

    'After an Undo or Redo call is invoked, we need to re-establish current non-destructive layer settings.
    ' (This allows us to detect changes to said settings, and create new Undo/Redo data accordingly.)
    Dim undoOrRedoUsed As Boolean

    If Strings.StringsEqual(processID, "Undo", True) Then
        
        If FormMain.MnuEdit(0).Enabled Then
            
            PDImages.GetActiveImage.UndoManager.RestoreUndoData
            Interface.NotifyImageChanged PDImages.GetActiveImageID()
            
            'Because Undo/Redo can involve image size changes (e.g. "Undo Resize Image"), we need to send a forcible
            ' UI notification to ensure that elements like rulers are correctly updated.
            Viewport.NotifyEveryoneOfViewportChanges
            
            undoOrRedoUsed = True
            
        End If
        Process_EditMenu = True
            
    ElseIf Strings.StringsEqual(processID, "Redo", True) Then
        If FormMain.MnuEdit(1).Enabled Then
            PDImages.GetActiveImage.UndoManager.RestoreRedoData
            Interface.NotifyImageChanged PDImages.GetActiveImageID()
            Viewport.NotifyEveryoneOfViewportChanges
            undoOrRedoUsed = True
        End If
        Process_EditMenu = True
        
    ElseIf Strings.StringsEqual(processID, "Undo history", True) Then
        If raiseDialog Then
            ShowPDDialog vbModal, FormUndoHistory
        Else
            PDImages.GetActiveImage.UndoManager.MoveToSpecificUndoPoint_XML processParameters
            Interface.NotifyImageChanged PDImages.GetActiveImageID()
            Viewport.NotifyEveryoneOfViewportChanges
            undoOrRedoUsed = True
        End If
        Process_EditMenu = True
        
    ElseIf Strings.StringsEqual(processID, "Fade", True) Then
        If raiseDialog Then ShowPDDialog vbModal, FormFadeLast Else FormFadeLast.fxFadeLastAction processParameters
        Process_EditMenu = True
        
    ElseIf Strings.StringsEqual(processID, "Cut", True) Then
        g_Clipboard.ClipboardCut False
        Process_EditMenu = True
        
    ElseIf Strings.StringsEqual(processID, "Cut merged", True) Then
        g_Clipboard.ClipboardCut True
        Process_EditMenu = True
        
    ElseIf Strings.StringsEqual(processID, "Copy", True) Then
        g_Clipboard.ClipboardCopy False
        Process_EditMenu = True
        
    ElseIf Strings.StringsEqual(processID, "Copy merged", True) Then
        g_Clipboard.ClipboardCopy True
        Process_EditMenu = True
    
```

---

</SwmSnippet>

<SwmSnippet path="/Modules/Processor.bas" line="1534">

---

After returning from dialog handling, `Process_EditMenu` wraps up by handling special clipboard and selection commands, and syncs layer properties if undo/redo was used.

```visual basic
    'Note the active image check; if no images are loaded, "Paste" gets silently rerouted to
    ' PD's "Paste to new image" handler.  (Note also that we deliberately do *not* pass process
    ' parameters to the function; those parameters contain cursor x/y position, if any - and if
    ' the paste function receives them, it will perform a "paste to cursor" op instead.)
    ElseIf Strings.StringsEqual(processID, "Paste", True) Then
        
        'Note if an image is active.  If one is *not* active, we will attempt to "paste as new image" instead
        Dim origState As Boolean: origState = PDImages.IsImageActive()
        
        'Perform the paste
        Dim pasteResult As Boolean: pasteResult = g_Clipboard.ClipboardPaste(PDImages.IsImageActive())
        
        'If an image is now loaded and 1) it wasn't originally, or 2) the paste failed, abandon Undo/Redo tagging
        If (PDImages.IsImageActive And ((Not origState) Or (Not pasteResult))) Then createUndo = UNDO_Nothing
        Process_EditMenu = True
    
    '"Paste to cursor" is identical to "paste", except we ensure process parameters get passed
    ' so the paste function can retrieve cursor position (and position the new layer accordingly)
    ElseIf Strings.StringsEqual(processID, "Paste to cursor", True) Then
        g_Clipboard.ClipboardPaste PDImages.IsImageActive(), , processParameters
        Process_EditMenu = True
        
    ElseIf Strings.StringsEqual(processID, "Paste to new image", True) Or Strings.StringsEqual(processID, "Paste as new image", True) Then
        g_Clipboard.ClipboardPaste False
        Process_EditMenu = True
    
    ElseIf Strings.StringsEqual(processID, "Cut special", True) Then
        If raiseDialog Then
            Dialogs.ShowClipboardDialog co_Cut
        Else
            g_Clipboard.ClipboardCutSpecial processParameters
        End If
        Process_EditMenu = True
        
    ElseIf Strings.StringsEqual(processID, "Copy special", True) Then
        If raiseDialog Then
            Dialogs.ShowClipboardDialog co_Copy
        Else
            g_Clipboard.ClipboardCopySpecial processParameters
        End If
        Process_EditMenu = True
        
    ElseIf Strings.StringsEqual(processID, "Paste special", True) Then
        If raiseDialog Then
            Dialogs.ShowClipboardDialog co_Paste
        Else
            'TODO
        End If
        Process_EditMenu = True
    
    ElseIf Strings.StringsEqual(processID, "Empty clipboard", True) Then
        g_Clipboard.ClipboardEmpty
        Process_EditMenu = True
    
    ElseIf Strings.StringsEqual(processID, "Clear", True) Then
        SelectionFilters.Selection_Clear raiseDialog
        Process_EditMenu = True
    
    ElseIf Strings.StringsEqual(processID, "Content-aware fill", True) Then
        SelectionFilters.Selection_ContentAwareFill raiseDialog, processParameters
        Process_EditMenu = True
    
    ElseIf Strings.StringsEqual(processID, "Fill", True) Then
        SelectionFilters.Selection_Fill raiseDialog, processParameters
        Process_EditMenu = True
        
    ElseIf Strings.StringsEqual(processID, "Stroke", True) Then
        SelectionFilters.Selection_Stroke raiseDialog, processParameters
        Process_EditMenu = True
        
    End If
    
    If undoOrRedoUsed Then
        
        'Synchronize any non-destructive settings to the currently active layer
        Processor.SyncAllGenericLayerProperties PDImages.GetActiveImage.GetActiveLayer
        Processor.SyncAllTextLayerProperties PDImages.GetActiveImage.GetActiveLayer
        
    End If
    
End Function
```

---

</SwmSnippet>

## Image Menu Command Routing

<SwmSnippet path="/Modules/Processor.bas" line="243">

---

Back in `Process`, after EditMenu, we check if the command was handled. If not, we pass it to Process_ImageMenu to see if it's an image command.

```visual basic
    'Image menu operations have been successfully migrated to XML strings.  (None of their functions raise special return conditions, FYI.)
    If (Not processFound) Then processFound = Process_ImageMenu(processID, raiseDialog, processParameters, createUndo, relevantTool, recordAction, returnDetails)
    
```

---

</SwmSnippet>

## Image Menu Command Dispatch and Dialog Routing

```mermaid
%%{init: {"flowchart": {"defaultRenderer": "elk"}} }%%
flowchart TD
    node1["User selects an image menu action"]
    click node1 openCode "Modules/Processor.bas:2204:2335"
    node1 --> node2{"Which action?"}
    click node2 openCode "Modules/Processor.bas:2209:2332"
    node2 -->|"Duplicate image"| node3["Duplicate the current image"]
    click node3 openCode "Modules/Processor.bas:2210:2211"
    node2 -->|"Resize/Content-aware/Canvas size"| node4{"Show dialog?"}
    click node4 openCode "Modules/Processor.bas:2216:2225"
    node4 -->|"Yes"| node5["Show relevant dialog"]
    click node5 openCode "Modules/DialogManager.bas:324:333"
    node4 -->|"No"| node6["Apply resize/canvas action directly"]
    click node6 openCode "Modules/Processor.bas:2216:2225"
    node2 -->|"Fit canvas/Crop/Trim/Straighten/Rotate/Flip"| node7{"Show dialog?"}
    click node7 openCode "Modules/Processor.bas:2227:2271"
    node7 -->|"Yes"| node8["Show relevant dialog"]
    click node8 openCode "Modules/DialogManager.bas:336:345"
    node7 -->|"No"| node9["Apply transformation directly"]
    click node9 openCode "Modules/Processor.bas:2227:2271"
    node2 -->|"Merge/Flatten layers"| node10{"Show dialog?"}
    click node10 openCode "Modules/Processor.bas:2274:2287"
    node10 -->|"Yes"| node11["Show flatten dialog if relevant"]
    click node11 openCode "Modules/Processor.bas:2283:2284"
    node10 -->|"No"| node12["Merge/flatten layers directly"]
    click node12 openCode "Modules/Processor.bas:2275:2286"
    node2 -->|"Animation/Color lookup/Compare"| node13{"Show dialog?"}
    click node13 openCode "Modules/Processor.bas:2290:2301"
    node13 -->|"Yes"| node14["Show relevant dialog"]
    click node14 openCode "Modules/Processor.bas:2291:2296"
    node13 -->|"No"| node15["Apply action directly"]
    click node15 openCode "Modules/Processor.bas:2292:2301"
    node2 -->|"Edit/Remove metadata"| node16{"Show dialog?"}
    click node16 openCode "Modules/Processor.bas:2303:2311"
    node16 -->|"Yes"| node17["Show metadata dialog"]
    click node17 openCode "Modules/Processor.bas:2306:2307"
    node16 -->|"No"| node18["Remove all metadata"]
    click node18 openCode "Modules/Processor.bas:2310:2311"
    node2 -->|"Count unique colors"| node19["Count unique colors"]
    click node19 openCode "Modules/Processor.bas:2314:2315"
    node2 -->|"Legacy/removed action"| node20["No operation (for compatibility)"]
    click node20 openCode "Modules/Processor.bas:2317:2332"
classDef HeadingStyle fill:#777777,stroke:#333,stroke-width:2px;
```

<SwmSnippet path="/Modules/Processor.bas" line="2204">

---

In `Process_ImageMenu`, we match processID to image commands. For those needing user input, we call dialog functions (like ShowResizeDialog); for others, we run the action directly. Legacy commands are handled as no-ops.

```visual basic
Private Function Process_ImageMenu(ByVal processID As String, Optional raiseDialog As Boolean = False, Optional processParameters As String = vbNullString, Optional createUndo As PD_UndoType = UNDO_Nothing, Optional relevantTool As Long = -1, Optional recordAction As Boolean = True, Optional ByRef returnDetails As String = vbNullString) As Boolean
    
    'It may seem odd, but the Duplicate function exists in the "Loading" module.  I do this because we effectively load a copy
    ' of the original image, so all loading operations (create pdImage object, catalog metadata, initialize properties) have to
    ' be repeated.
    If Strings.StringsEqual(processID, "Duplicate image", True) Then
        Loading.DuplicateCurrentImage
        Process_ImageMenu = True
    
    'Resize operations; note that prior to 6.4, "Resize" was used in place of "Resize image".  To preserve functionality of old macros,
    ' we add the old "Resize" operator here as well.
    ElseIf Strings.StringsEqual(processID, "Resize image", True) Or Strings.StringsEqual(processID, "Resize", True) Then
        If raiseDialog Then ShowResizeDialog pdat_Image Else FormResize.ResizeImage processParameters
        Process_ImageMenu = True
        
    ElseIf Strings.StringsEqual(processID, "Content-aware image resize", True) Then
```

---

</SwmSnippet>

<SwmSnippet path="/Modules/DialogManager.bas" line="324">

---

`ShowResizeDialog` sets up the target for resizing, then calls ShowPDDialog to display the resize dialog modally.

```visual basic
Public Sub ShowResizeDialog(ByVal ResizeTarget As PD_ActionTarget)
    FormResize.ResizeTarget = ResizeTarget
    ShowPDDialog vbModal, FormResize
End Sub
```

---

</SwmSnippet>

<SwmSnippet path="/Modules/Processor.bas" line="2220">

---

Back in `Process_ImageMenu`, after handling resize, we use the same dialog manager pattern for content-aware resizing—show the dialog if needed, otherwise run the action directly.

```visual basic
        If raiseDialog Then ShowContentAwareResizeDialog pdat_Image Else FormResizeContentAware.SmartResizeImage processParameters
        Process_ImageMenu = True
        
    ElseIf Strings.StringsEqual(processID, "Canvas size", True) Then
```

---

</SwmSnippet>

<SwmSnippet path="/Modules/DialogManager.bas" line="330">

---

`ShowContentAwareResizeDialog` just sets the resize target and calls ShowPDDialog to show the dialog. That's it.

```visual basic
Public Sub ShowContentAwareResizeDialog(ByVal ResizeTarget As PD_ActionTarget)
    FormResizeContentAware.ResizeTarget = ResizeTarget
    ShowPDDialog vbModal, FormResizeContentAware
End Sub
```

---

</SwmSnippet>

<SwmSnippet path="/Modules/Processor.bas" line="2224">

---

Back in `Process_ImageMenu`, for canvas size, we either show the dialog (ShowPDDialog) or run the resize directly, depending on raiseDialog.

```visual basic
        If raiseDialog Then ShowPDDialog vbModal, FormCanvasSize Else FormCanvasSize.ResizeCanvas processParameters
        Process_ImageMenu = True
            
    ElseIf Strings.StringsEqual(processID, "Fit canvas to active layer", True) Then
        Filters_Transform.FitCanvasToLayer_XML processParameters
        Process_ImageMenu = True
        
    ElseIf Strings.StringsEqual(processID, "Fit canvas around all layers", True) Then
        Filters_Transform.MenuFitCanvasToAllLayers
        Process_ImageMenu = True
    
```

---

</SwmSnippet>

<SwmSnippet path="/Modules/Processor.bas" line="2235">

---

Back in `Process_ImageMenu`, for crop and straighten, we call dialog manager functions to show dialogs if raiseDialog is True, or run the action directly otherwise.

```visual basic
    'Crop operations.  Note that the main form submits "Crop" requests with raiseDialog set to TRUE.  This tells us to ask the
    ' crop handler if a non-destructive crop is possible.  It will then submit a second "Crop" requests with raiseDialog set to FALSE.
    ElseIf Strings.StringsEqual(processID, "Crop", True) Then
        If raiseDialog Then Filters_Transform.SeeIfCropCanBeAppliedNonDestructively Else Filters_Transform.CropToSelection_XML processParameters
        Process_ImageMenu = True
        
    ElseIf Strings.StringsEqual(processID, "Trim empty image borders", True) Then
        Filters_Transform.TrimImage
        Process_ImageMenu = True
        
    ElseIf Strings.StringsEqual(processID, "Straighten image", True) Then
        If raiseDialog Then ShowStraightenDialog pdat_Image Else FormStraighten.StraightenImage processParameters
        Process_ImageMenu = True
            
    ElseIf Strings.StringsEqual(processID, "Rotate image 90 clockwise", True) Or Strings.StringsEqual(processID, "Rotate 90 clockwise", True) Then
        Filters_Transform.MenuRotate90Clockwise
        Process_ImageMenu = True
            
    ElseIf Strings.StringsEqual(processID, "Rotate image 180", True) Or Strings.StringsEqual(processID, "Rotate 180", True) Then
        Filters_Transform.MenuRotate180
        Process_ImageMenu = True
            
    ElseIf Strings.StringsEqual(processID, "Rotate image 90 counter-clockwise", True) Or Strings.StringsEqual(processID, "Rotate 90 counter-clockwise", True) Then
        Filters_Transform.MenuRotate270Clockwise
        Process_ImageMenu = True
            
    ElseIf Strings.StringsEqual(processID, "Arbitrary image rotation", True) Or Strings.StringsEqual(processID, "Arbitrary rotation", True) Then
```

---

</SwmSnippet>

<SwmSnippet path="/Modules/DialogManager.bas" line="342">

---

`ShowStraightenDialog` sets the target for straightening, then shows the dialog modally so the user has to interact before continuing.

```visual basic
Public Sub ShowStraightenDialog(ByVal StraightenTarget As PD_ActionTarget)
    FormStraighten.StraightenTarget = StraightenTarget
    ShowPDDialog vbModal, FormStraighten
End Sub
```

---

</SwmSnippet>

<SwmSnippet path="/Modules/Processor.bas" line="2262">

---

Back in `Process_ImageMenu`, for rotation, we call dialog manager functions to show dialogs if raiseDialog is True, or run the action directly otherwise.

```visual basic
        If raiseDialog Then ShowRotateDialog pdat_Image Else FormRotate.RotateArbitrary processParameters
        Process_ImageMenu = True
            
    ElseIf Strings.StringsEqual(processID, "Flip image vertically", True) Or Strings.StringsEqual(processID, "Flip vertically", True) Then
        Filters_Transform.MenuFlip
        Process_ImageMenu = True
            
    ElseIf Strings.StringsEqual(processID, "Flip image horizontally", True) Or Strings.StringsEqual(processID, "Flip horizontally", True) Then
        Filters_Transform.MenuMirror
        Process_ImageMenu = True
    
```

---

</SwmSnippet>

<SwmSnippet path="/Modules/DialogManager.bas" line="336">

---

`ShowRotateDialog` sets the target for rotation, then shows the dialog modally so the user has to interact before continuing.

```visual basic
Public Sub ShowRotateDialog(ByVal RotateTarget As PD_ActionTarget)
    FormRotate.RotateTarget = RotateTarget
    ShowPDDialog vbModal, FormRotate
End Sub
```

---

</SwmSnippet>

<SwmSnippet path="/Modules/Processor.bas" line="2273">

---

Back in `Process_ImageMenu`, for flatten and animation options, we show dialogs if needed (ShowPDDialog) or run the action directly, depending on raiseDialog.

```visual basic
    'Merge visible layers
    ElseIf Strings.StringsEqual(processID, "Merge visible layers", True) Then
        Layers.MergeVisibleLayers
        Process_ImageMenu = True
        
    'Flatten image.  This dialog is a little weird because we don't *always* show it.  If an image has
    ' no transparency, we don't need to prompt for transparency handling - so we always check state in
    ' advance, rather than bother the user with an unnecessary prompt.
    ElseIf Strings.StringsEqual(processID, "Flatten image", True) Then
        If raiseDialog Then
            If Layers.IsFlattenDialogRelevant() Then ShowPDDialog vbModal, FormLayerFlatten Else Processor.Process "Flatten image", False, vbNullString, UNDO_Image
        Else
            Layers.FlattenImage processParameters
        End If
        Process_ImageMenu = True
    
    'Modify animation settings
    ElseIf Strings.StringsEqual(processID, "Animation options", True) Then
        If raiseDialog Then ShowPDDialog vbModal, FormAnimation Else FormAnimation.ApplyAnimationChanges processParameters
        Process_ImageMenu = True
    
    'Compare two images/layers
    ElseIf Strings.StringsEqual(processID, "Create color lookup", True) Then
        If raiseDialog Then ShowPDDialog vbModal, FormImageCreateLUT Else FormImageCreateLUT.CreateDifferenceLUT processParameters
        Process_ImageMenu = True
        
    ElseIf Strings.StringsEqual(processID, "Compare similarity", True) Then
        If raiseDialog Then ShowPDDialog vbModal, FormImageCompare Else FormImageCompare.CompareImages processParameters
        Process_ImageMenu = True
        
    ElseIf Strings.StringsEqual(processID, "Edit metadata", True) Then
        
```

---

</SwmSnippet>

<SwmSnippet path="/Modules/Processor.bas" line="2305">

---

After returning from dialog handling, `Process_ImageMenu` wraps up by handling metadata, color counting, and legacy commands as no-ops to keep macro compatibility.

```visual basic
        'Note that there is no "Else" block here; the "Else" block does nothing but notify the processor to create an Undo entry
        If raiseDialog Then ExifTool.ShowMetadataDialog PDImages.GetActiveImage()
        Process_ImageMenu = True
        
    ElseIf Strings.StringsEqual(processID, "Remove all metadata", True) Then
        ExifTool.RemoveAllMetadata PDImages.GetActiveImage()
        Process_ImageMenu = True
        
    ElseIf Strings.StringsEqual(processID, "Count unique colors", True) Then
        Filters_Miscellaneous.MenuCountColors
        Process_ImageMenu = True
        
    'NOTE!  Some Image-menu actions have been removed in new versions of the programs.  If they exist inside macros,
    ' I don't want to raise errors, so I've included their keywords here even though they are basically NOPs.
    
    'TODO 8.2: reinstate auto-cropping
    ElseIf Strings.StringsEqual(processID, "Autocrop", True) Then
    '    AutocropImage
        Process_ImageMenu = True
    
    'Isometric conversion was removed in v6.4.  There are not currently plans to reinstate it.
    ElseIf Strings.StringsEqual(processID, "Isometric conversion", True) Then
        Process_ImageMenu = True
    
    'Image > Tile was removed in v7.0.  There are not currently plans to reinstate it.
    ElseIf Strings.StringsEqual(processID, "Tile", True) Then
        Process_ImageMenu = True
    
    End If
       
End Function
```

---

</SwmSnippet>

## Layer Menu Command Routing

<SwmSnippet path="/Modules/Processor.bas" line="246">

---

Back in `Process`, after ImageMenu, we check if the command was handled. If not, we pass it to Process_LayerMenu to see if it's a layer command.

```visual basic
    'Layer menu operations have been successfully migrated to XML strings.  (None of their functions raise special return conditions, FYI.)
    If (Not processFound) Then processFound = Process_LayerMenu(processID, raiseDialog, processParameters, createUndo, relevantTool, recordAction, returnDetails)
    
```

---

</SwmSnippet>

## Layer Command Dispatch and Execution

```mermaid
%%{init: {"flowchart": {"defaultRenderer": "elk"}} }%%
flowchart TD
    node1["Receive layer operation request"] --> node2{"What type of layer action?"}
    click node1 openCode "Modules/Processor.bas:2340:2628"
    node2 -->|"Add layer"| node3{"Is this a text layer and macro/batch is active?"}
    node2 -->|"Remove layer"| node4["Remove the selected layer(s)"]
    node2 -->|"Duplicate layer"| node5["Duplicate the selected layer"]
    node2 -->|"Merge/move/transform layer"| node6{"Show dialog to user?"}
    node2 -->|"Change visibility"| node7["Show or hide layers"]
    node2 -->|"Resize/crop/alpha"| node8{"Show dialog to user?"}
    node2 -->|"Split/rasterize/rearrange"| node9["Split, rasterize, or rearrange layers"]
    click node2 openCode "Modules/Processor.bas:2348:2626"
    node3 -->|"Yes"| node10["Create and initialize text layer from parameters"]
    click node3 openCode "Modules/Processor.bas:2361:2382"
    node3 -->|"No"| node11["Add a new blank or standard layer"]
    click node11 openCode "Modules/Processor.bas:2349:2355"
    node6 -->|"Yes"| node12["Show dialog for user input"]
    node6 -->|"No"| node13["Perform action directly"]
    node8 -->|"Yes"| node14["Show dialog for user input"]
    node8 -->|"No"| node15["Perform action directly"]
    node4 --> node16["Operation complete"]
    click node4 openCode "Modules/Processor.bas:2405:2412"
    node5 --> node16
    click node5 openCode "Modules/Processor.bas:2388:2391"
    node7 --> node16
    click node7 openCode "Modules/Processor.bas:2476:2497"
    node9 --> node16
    click node9 openCode "Modules/Processor.bas:2581:2626"
    node10 --> node16
    click node10 openCode "Modules/Processor.bas:2363:2382"
    node11 --> node16
    node12 --> node16
    node13 --> node16
    node14 --> node16
    node15 --> node16
    click node16 openCode "Modules/Processor.bas:2351:2626"
classDef HeadingStyle fill:#777777,stroke:#333,stroke-width:2px;
```

<SwmSnippet path="/Modules/Processor.bas" line="2340">

---

In Process_LayerMenu, we match the incoming processID against a big list of known layer commands. For each match, we either show a dialog (if raiseDialog is set) or call the relevant Layers method directly, passing parameters parsed from XML. This setup means we can handle both interactive and automated actions, and the XML parsing lets us pass around complex data for things like multi-layer ops. We call into Interface.bas next when a dialog is needed, so the user can provide input before the action runs.

```visual basic
Private Function Process_LayerMenu(ByVal processID As String, Optional ByVal raiseDialog As Boolean = False, Optional ByRef processParameters As String = vbNullString, Optional ByRef createUndo As PD_UndoType = UNDO_Nothing, Optional ByVal relevantTool As Long = -1, Optional ByRef recordAction As Boolean = True, Optional ByRef returnDetails As String = vbNullString) As Boolean
    
    'A number of layer functions pass the relevant layer index in the parameter string (as future-proofing against selecting
    ' multiple layers).  To simplify the parsing of these entries, we always create an XML parser.
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    cParams.SetParamString processParameters
    
    'Add layers to an image
    If Strings.StringsEqual(processID, "Add blank layer", True) Then
        Layers.AddBlankLayer_XML processParameters
        Process_LayerMenu = True
    
    ElseIf Strings.StringsEqual(processID, "Add new layer", True) Then
        If raiseDialog Then ShowPDDialog vbModal, FormNewLayer Else Layers.AddNewLayer_XML processParameters
        Process_LayerMenu = True
    
```

---

</SwmSnippet>

<SwmSnippet path="/Modules/Processor.bas" line="2357">

---

After returning from Interface.bas (if a dialog was shown), Process_LayerMenu handles text layer commands. If we're running a macro or batch, it creates the text layer and restores its state from XML, so automated tasks can recreate layers exactly. Otherwise, it just updates undo/redo for normal usage. DialogManager.bas is called next for any actions that need user input via dialogs.

```visual basic
    'During normal usage, "New text layer" is a dummy entry used by the on-canvas text tool.  It is called *after* a new layer
    ' has already been created, and the sole purpose of the function is to add the newly created text layer to the Undo/Redo chain.
    '
    'During macro playback, "New text layer" actually means *create* a new text layer, using the settings specified in the parameter string.
    ElseIf Strings.StringsEqual(processID, "New text layer", True) Or Strings.StringsEqual(processID, "New typography layer", True) Then
        
        If ((Macros.GetMacroStatus = MacroPLAYBACK) Or (Macros.GetMacroStatus = MacroBATCH)) Then
            
            'Start by creating a new layer
            If Strings.StringsEqual(processID, "New text layer", True) Then
                Layers.AddNewLayer PDImages.GetActiveImage.GetActiveLayerIndex, PDL_TextBasic, 0, 0, 0, True, vbNullString, 0#, 0#, True
            Else
                Layers.AddNewLayer PDImages.GetActiveImage.GetActiveLayerIndex, PDL_TextAdvanced, 0, 0, 0, True, vbNullString, 0#, 0#, True
            End If
            
            'Text layer parameters can be precisely recreated in two steps:
            
            '1) Initialize the standard layer header
            PDImages.GetActiveImage.GetActiveLayer.CreateNewLayerFromXML cParams.GetString("layerheader")
            
            '2) Initialize the text-layer-specific bits
            PDImages.GetActiveImage.GetActiveLayer.SetVectorDataFromXML cParams.GetString("layerdata")
            
        End If
        
        Process_LayerMenu = True
    
    ElseIf Strings.StringsEqual(processID, "New layer from file", True) Then
        Layers.LoadImageAsNewLayer raiseDialog, processParameters
        Process_LayerMenu = True
    
    ElseIf Strings.StringsEqual(processID, "Duplicate layer", True) Then
        Layers.DuplicateLayerByIndex_XML processParameters
        Process_LayerMenu = True
        
    ElseIf Strings.StringsEqual(processID, "New layer from visible layers", True) Then
        Layers.AddLayerFromVisibleLayers
        Process_LayerMenu = True
        
    ElseIf Strings.StringsEqual(processID, "Layer via copy", True) Then
        Layers.AddLayerViaCopy
        Process_LayerMenu = True
        
    ElseIf Strings.StringsEqual(processID, "Layer via cut", True) Then
        Layers.AddLayerViaCut
        Process_LayerMenu = True
        
    'Remove layers from an image
    ElseIf Strings.StringsEqual(processID, "Delete layer", True) Then
        Layers.DeleteLayer_XML processParameters
        Process_LayerMenu = True
    
    ElseIf Strings.StringsEqual(processID, "Delete hidden layers", True) Then
        Layers.DeleteHiddenLayers
        Process_LayerMenu = True
    
    'Replace layer contents with something new
    ElseIf Strings.StringsEqual(processID, "Replace layer from clipboard", True) Then
        If (Not Layers.ReplaceLayerWithClipboard) Then createUndo = UNDO_Nothing
        Process_LayerMenu = True
    
    ElseIf Strings.StringsEqual(processID, "Replace layer from file", True) Then
        Layers.LoadImageAsNewLayer raiseDialog, processParameters, replaceActiveLayerInstead:=True
        Process_LayerMenu = True
    
    ElseIf Strings.StringsEqual(processID, "Replace layer from visible layers", True) Then
        Layers.AddLayerFromVisibleLayers True
        Process_LayerMenu = True
    
    'Merge a layer up or down
    ElseIf Strings.StringsEqual(processID, "Merge layer down", True) Then
        Layers.MergeLayerAdjacent cParams.GetLong("layerindex"), True
        Process_LayerMenu = True
        
    ElseIf Strings.StringsEqual(processID, "Merge layer up", True) Then
        Layers.MergeLayerAdjacent cParams.GetLong("layerindex"), False
        Process_LayerMenu = True
    
    'Select top/up/below/bottom layer
    ElseIf Strings.StringsEqual(processID, "Go to top layer", True) Then
        Layers.SelectLayerTopBottom True
        Process_LayerMenu = True
    
    ElseIf Strings.StringsEqual(processID, "Go to layer above", True) Then
        Layers.SelectLayerAdjacent True
        Process_LayerMenu = True
    
    ElseIf Strings.StringsEqual(processID, "Go to layer below", True) Then
        Layers.SelectLayerAdjacent False
        Process_LayerMenu = True
    
    ElseIf Strings.StringsEqual(processID, "Go to bottom layer", True) Then
        Layers.SelectLayerTopBottom False
        Process_LayerMenu = True
    
    'Raise a layer up or down
    ElseIf Strings.StringsEqual(processID, "Raise layer", True) Then
        Layers.MoveLayerAdjacent cParams.GetLong("layerindex"), True
        Process_LayerMenu = True
    
    ElseIf Strings.StringsEqual(processID, "Lower layer", True) Then
        Layers.MoveLayerAdjacent cParams.GetLong("layerindex"), False
        Process_LayerMenu = True
        
    'Raise or lower to layer to end of stack
    ElseIf Strings.StringsEqual(processID, "Raise layer to top", True) Then
        Layers.MoveLayerToEndOfStack cParams.GetLong("layerindex"), True
        Process_LayerMenu = True
    
    ElseIf Strings.StringsEqual(processID, "Lower layer to bottom", True) Then
        Layers.MoveLayerToEndOfStack cParams.GetLong("layerindex"), False
        Process_LayerMenu = True
        
    'Reverse layer order
    ElseIf Strings.StringsEqual(processID, "Reverse layer order", True) Then
        Layers.ReverseLayerOrder
        Process_LayerMenu = True
    
    'Toggle active layer visibility
    ElseIf Strings.StringsEqual(processID, "Toggle layer visibility", True) Then
        Layers.ToggleLayerVisibility cParams.GetLong("layerindex")
        Process_LayerMenu = True
    
    'Show or hide just the active layer
    ElseIf Strings.StringsEqual(processID, "Show only this layer", True) Then
        Layers.MakeJustOneLayerVisible cParams.GetLong("layerindex", PDImages.GetActiveImage.GetActiveLayerIndex)
        Process_LayerMenu = True
        
    ElseIf Strings.StringsEqual(processID, "Hide only this layer", True) Then
        Layers.MakeJustOneLayerHidden cParams.GetLong("layerindex", PDImages.GetActiveImage.GetActiveLayerIndex)
        Process_LayerMenu = True
    
    'Show or hide all layers
    ElseIf Strings.StringsEqual(processID, "Show all layers", True) Then
        Layers.SetLayerVisibility_AllLayers True
        Process_LayerMenu = True
        
    ElseIf Strings.StringsEqual(processID, "Hide all layers", True) Then
        Layers.SetLayerVisibility_AllLayers False
        Process_LayerMenu = True
    
    'Crop tasks
    ElseIf Strings.StringsEqual(processID, "Crop layer to selection", True) Then
        Filters_Transform.CropToSelection PDImages.GetActiveImage.GetActiveLayerIndex
        Process_LayerMenu = True
        
    ElseIf Strings.StringsEqual(processID, "Pad layer to image size", True) Then
        Layers.PadToImageSize PDImages.GetActiveImage, PDImages.GetActiveImage.GetActiveLayerIndex
        Process_LayerMenu = True
    
    ElseIf Strings.StringsEqual(processID, "Trim empty layer borders", True) Then
        Layers.TrimEmptyBorders PDImages.GetActiveImage, PDImages.GetActiveImage.GetActiveLayerIndex
        Process_LayerMenu = True
    
    'Non-destructive layer size and orientation changes
    ElseIf Strings.StringsEqual(processID, "Reset layer size", True) Then
        Layers.ResetLayerSize cParams.GetLong("layerindex")
        Process_LayerMenu = True
    
    ' (Just kidding, this action is destructive, but it sits on the non-destructive panel so I've included it here)
    ElseIf Strings.StringsEqual(processID, "Make layer changes permanent", True) Then
        Layers.MakeLayerAffineTransformsPermanent cParams.GetLong("layerindex")
        Process_LayerMenu = True
        
    ElseIf Strings.StringsEqual(processID, "Fit layer to image", True) Then
        Layers.FitLayerToImageSize cParams.GetLong("layerindex")
        Process_LayerMenu = True
        
    'Destructive layer orientation changes
    ElseIf Strings.StringsEqual(processID, "Straighten layer", True) Then
        If raiseDialog Then ShowStraightenDialog pdat_SingleLayer Else FormStraighten.StraightenImage processParameters
        Process_LayerMenu = True
        
    ElseIf Strings.StringsEqual(processID, "Rotate layer 90 clockwise", True) Then
        Filters_Transform.MenuRotate90Clockwise PDImages.GetActiveImage.GetActiveLayerIndex
        Process_LayerMenu = True
        
    ElseIf Strings.StringsEqual(processID, "Rotate layer 180", True) Then
        Filters_Transform.MenuRotate180 PDImages.GetActiveImage.GetActiveLayerIndex
        Process_LayerMenu = True
        
    ElseIf Strings.StringsEqual(processID, "Rotate layer 90 counter-clockwise", True) Then
        Filters_Transform.MenuRotate270Clockwise PDImages.GetActiveImage.GetActiveLayerIndex
        Process_LayerMenu = True
        
    ElseIf Strings.StringsEqual(processID, "Arbitrary layer rotation", True) Then
```

---

</SwmSnippet>

<SwmSnippet path="/Modules/Processor.bas" line="2543">

---

After returning from DialogManager.bas (if a dialog was shown), Process_LayerMenu keeps routing layer actions that need user input through the right dialog manager function. For example, arbitrary rotation and resizing each get their own dialog if raiseDialog is set. This keeps the UI consistent and lets users tweak settings before running the action.

```visual basic
        If raiseDialog Then ShowRotateDialog pdat_SingleLayer Else FormRotate.RotateArbitrary processParameters
        Process_LayerMenu = True
        
    ElseIf Strings.StringsEqual(processID, "Flip layer horizontally", True) Then
        Filters_Transform.MenuMirror PDImages.GetActiveImage.GetActiveLayerIndex
        Process_LayerMenu = True
    
    ElseIf Strings.StringsEqual(processID, "Flip layer vertically", True) Then
        Filters_Transform.MenuFlip PDImages.GetActiveImage.GetActiveLayerIndex
        Process_LayerMenu = True
            
```

---

</SwmSnippet>

<SwmSnippet path="/Modules/Processor.bas" line="2554">

---

After returning from DialogManager.bas, Process_LayerMenu checks if the next action (like resize or content-aware resize) needs a dialog. If so, it shows the right one; otherwise, it just runs the effect. This keeps things flexible for both interactive and automated use.

```visual basic
    'Destructive layer size changes
    ElseIf Strings.StringsEqual(processID, "Resize layer", True) Then
        If raiseDialog Then ShowResizeDialog pdat_SingleLayer Else FormResize.ResizeImage processParameters
        Process_LayerMenu = True
        
    ElseIf Strings.StringsEqual(processID, "Content-aware layer resize", True) Then
```

---

</SwmSnippet>

<SwmSnippet path="/Modules/Processor.bas" line="2560">

---

After returning from DialogManager.bas, Process_LayerMenu checks if content-aware layer resize needs a dialog. If so, it shows it; otherwise, it just runs the resize. This keeps the workflow consistent for both manual and automated actions.

```visual basic
        If raiseDialog Then ShowContentAwareResizeDialog pdat_SingleLayer Else FormResizeContentAware.SmartResizeImage processParameters
        Process_LayerMenu = True
        
```

---

</SwmSnippet>

<SwmSnippet path="/Modules/Processor.bas" line="2563">

---

After returning from DialogManager.bas, Process_LayerMenu checks if the next alpha-related action needs a dialog. If so, it shows the right form from Interface.bas; otherwise, it just runs the effect. This lets users interact when needed, or automates the process if not.

```visual basic
    'Change layer alpha
    ElseIf Strings.StringsEqual(processID, "Color to alpha", True) Then
        If raiseDialog Then ShowPDDialog vbModal, FormTransparency_FromColor Else FormTransparency_FromColor.ColorToAlpha processParameters
        Process_LayerMenu = True
        
    ElseIf Strings.StringsEqual(processID, "Luminance to alpha", True) Then
        If raiseDialog Then ShowPDDialog vbModal, FormTransparency_FromLuma Else FormTransparency_FromLuma.LuminanceToAlpha processParameters
        Process_LayerMenu = True
        
    ElseIf Strings.StringsEqual(processID, "Remove alpha channel", True) Then
        If raiseDialog Then ShowPDDialog vbModal, FormConvert24bpp Else FormConvert24bpp.RemoveLayerTransparency processParameters
        Process_LayerMenu = True
    
    ElseIf Strings.StringsEqual(processID, "Threshold alpha", True) Then
        If raiseDialog Then ShowPDDialog vbModal, FormThresholdAlpha Else FormThresholdAlpha.FxThresholdAlpha processParameters
        Process_LayerMenu = True
    
    'Convert layers to images (or images to layers)
    ElseIf Strings.StringsEqual(processID, "Split layer into image", True) Then
        Layers.SplitLayerToImage BuildParamList("target-layer", PDImages.GetActiveImage.GetActiveLayerIndex)
        Process_LayerMenu = True
        
    ElseIf Strings.StringsEqual(processID, "Split layers into images", True) Then
        Layers.SplitLayerToImage BuildParamList("target-layer", -1)
        Process_LayerMenu = True
        
    ElseIf Strings.StringsEqual(processID, "Split images into layers", True) Then
        If raiseDialog Then ShowPDDialog vbModal, FormLayerSplit Else Layers.MergeImagesToLayers processParameters
        Process_LayerMenu = True
        
```

---

</SwmSnippet>

<SwmSnippet path="/Modules/Processor.bas" line="2593">

---

After returning from Interface.bas, Process_LayerMenu finishes up by handling the rest of the layer commands—rasterizing, on-canvas moves, and dummy entries for undo/redo. The function centralizes all layer-related actions, making it easier to manage, and hooks into undo/redo and macro recording as needed.

```visual basic
    'Rasterizing
    ElseIf Strings.StringsEqual(processID, "Rasterize layer", True) Then
        Layers.RasterizeLayer cParams.GetLong("layerindex")
        Process_LayerMenu = True
    
    ElseIf Strings.StringsEqual(processID, "Rasterize all layers", True) Then
        Layers.RasterizeLayer -1
        Process_LayerMenu = True
    
    'On-canvas layer modifications (moving, non-destructive resizing, etc)
    ElseIf Strings.StringsEqual(processID, "Resize layer (on-canvas)", True) Then
        Layers.ResizeLayerNonDestructive PDImages.GetActiveImage.GetActiveLayerIndex, processParameters
        Process_LayerMenu = True
    
    ElseIf Strings.StringsEqual(processID, "Rotate layer (on-canvas)", True) Then
        Layers.RotateLayerNonDestructive PDImages.GetActiveImage.GetActiveLayerIndex, processParameters
        Process_LayerMenu = True
    
    ElseIf Strings.StringsEqual(processID, "Move layer", True) Then
        Layers.MoveLayerOnCanvas PDImages.GetActiveImage.GetActiveLayerIndex, processParameters
        Process_LayerMenu = True
    
    'If a selection is active, the user can use the move tool to copy (or cut) just the selected
    ' pixels from either the active layer or the full image stack.  When this operation occurs,
    ' we use a different processor call to make the Undo/Redo op title more sensible.
    ElseIf Strings.StringsEqual(processID, "Move selected pixels", True) Then
        Layers.MoveLayerOnCanvas PDImages.GetActiveImage.GetActiveLayerIndex, processParameters
        Process_LayerMenu = True
    
    '"Rearrange layers" is a dummy entry.  It does not actually modify the image; its sole purpose is
    ' to create an Undo/Redo entry after the user has performed a drag/drop rearrangement of the layer stack.
    ElseIf Strings.StringsEqual(processID, "Rearrange layers", True) Then
        Process_LayerMenu = True
    End If
    
End Function
```

---

</SwmSnippet>

## Selection Command Routing

```mermaid
%%{init: {"flowchart": {"defaultRenderer": "elk"}} }%%
flowchart TD
  node1["Start: Has the operation already been handled?"]
  click node1 openCode "Modules/Processor.bas:249:249"
  node1 --> node2{"Operation already handled?"}
  click node2 openCode "Modules/Processor.bas:249:249"
  node2 -->|"No"| node3["Try selection menu operation"]
  click node3 openCode "Modules/Processor.bas:249:250"
  node3 --> node4{"Operation handled by selection menu?"}
  click node4 openCode "Modules/Processor.bas:253:253"
  node4 -->|"No"| node5["Try adjustment menu operation"]
  click node5 openCode "Modules/Processor.bas:253:253"
  node4 -->|"Yes"| node6["Operation handled"]
  node2 -->|"Yes"| node6["Operation handled"]
  click node6 openCode "Modules/Processor.bas:249:253"

classDef HeadingStyle fill:#777777,stroke:#333,stroke-width:2px;
```

<SwmSnippet path="/Modules/Processor.bas" line="249">

---

After returning from Process_LayerMenu, Process checks if the command was handled. If not, it passes the processID to Process_SelectMenu, so selection-related commands get a shot at handling it next. This keeps the routing clean and predictable.

```visual basic
    'Select menu operations have been successfully migrated to XML strings.  (None of their functions raise special return conditions, FYI.)
    If (Not processFound) Then processFound = Process_SelectMenu(processID, raiseDialog, processParameters, createUndo, relevantTool, recordAction, returnDetails)
    
```

---

</SwmSnippet>

<SwmSnippet path="/Modules/Processor.bas" line="2633">

---

Process_SelectMenu checks the processID and runs the right selection operation, like create, remove, invert, or modify. It uses XML for parameters so we can handle more complex cases (like multi-layer selections) and supports both dialogs and direct execution. Dummy entries are included for undo/redo tracking when selections are moved or resized.

```visual basic
Private Function Process_SelectMenu(ByVal processID As String, Optional raiseDialog As Boolean = False, Optional processParameters As String = vbNullString, Optional createUndo As PD_UndoType = UNDO_Nothing, Optional relevantTool As Long = -1, Optional recordAction As Boolean = True, Optional ByRef returnDetails As String = vbNullString) As Boolean
    
    'A number of selection functions pass the relevant layer index in the parameter string (as future-proofing against selecting
    ' multiple layers).  To simplify the parsing of these entries, we always create an XML parser.
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    cParams.SetParamString processParameters
        
    'Create/remove selections
    If Strings.StringsEqual(processID, "Create selection", True) Then
        Selections.CreateNewSelection processParameters
        Process_SelectMenu = True
    
    ElseIf Strings.StringsEqual(processID, "Remove selection", True) Then
        Selections.RemoveCurrentSelection
        Process_SelectMenu = True
        
    'Modify the existing selection in some way
    ElseIf Strings.StringsEqual(processID, "Invert selection", True) Then
        SelectionFilters.Selection_Invert
        Process_SelectMenu = True
        
    ElseIf Strings.StringsEqual(processID, "Grow selection", True) Then
        If raiseDialog Then SelectionFilters.Selection_Grow True Else SelectionFilters.Selection_Grow False, cParams.GetDouble("filtervalue")
        Process_SelectMenu = True
        
    ElseIf Strings.StringsEqual(processID, "Shrink selection", True) Then
        If raiseDialog Then SelectionFilters.Selection_Shrink True Else SelectionFilters.Selection_Shrink False, cParams.GetDouble("filtervalue")
        Process_SelectMenu = True
    
    ElseIf Strings.StringsEqual(processID, "Feather selection", True) Then
        If raiseDialog Then SelectionFilters.Selection_Blur True Else SelectionFilters.Selection_Blur False, cParams.GetDouble("filtervalue")
        Process_SelectMenu = True
        
    ElseIf Strings.StringsEqual(processID, "Sharpen selection", True) Then
        If raiseDialog Then SelectionFilters.Selection_Sharpen True Else SelectionFilters.Selection_Sharpen False, cParams.GetDouble("filtervalue")
        Process_SelectMenu = True
        
    ElseIf Strings.StringsEqual(processID, "Border selection", True) Then
        If raiseDialog Then SelectionFilters.Selection_ConvertToBorder True Else SelectionFilters.Selection_ConvertToBorder False, cParams.GetDouble("filtervalue")
        Process_SelectMenu = True
    
    'Modify selected pixels in various ways
    ElseIf Strings.StringsEqual(processID, "Erase selected area", True) Then
        Selections.EraseSelectedArea cParams.GetLong("targetlayer")
        Process_SelectMenu = True
    
    ElseIf Strings.StringsEqual(processID, "Fill selected area", True) Then
        SelectionFilters.Selection_Fill raiseDialog, processParameters
        Process_SelectMenu = True
        
    ElseIf Strings.StringsEqual(processID, "Heal selected area", True) Then
        SelectionFilters.Selection_ContentAwareFill raiseDialog, processParameters
        Process_SelectMenu = True
        
    ElseIf Strings.StringsEqual(processID, "Stroke selection outline", True) Then
        SelectionFilters.Selection_Stroke raiseDialog, processParameters
        Process_SelectMenu = True
        
    'Load/save selection from/to file
    ElseIf Strings.StringsEqual(processID, "Load selection", True) Then
        If raiseDialog Then SelectionFiles.LoadSelectionFromFile True Else SelectionFiles.LoadSelectionFromFile False, processParameters
        Process_SelectMenu = True
        
    ElseIf Strings.StringsEqual(processID, "Save selection", True) Then
        SelectionFiles.SaveSelectionToFile
        Process_SelectMenu = True
        
    'Export selected area as image (defaults to PNG, but user can select the actual format)
    ElseIf Strings.StringsEqual(processID, "Export selected area as image", True) Then
        SelectionFiles.ExportSelectedAreaAsImage
        Process_SelectMenu = True
    
    'Export selection mask as image (defaults to PNG, but user can select the actual format)
    ElseIf Strings.StringsEqual(processID, "Export selection mask as image", True) Then
        SelectionFiles.ExportSelectionMaskAsImage
        Process_SelectMenu = True
    
    ' This is a dummy entry; it only exists so that Undo/Redo data is correctly generated when a selection is moved
    ElseIf Strings.StringsEqual(processID, "Move selection", True) Then
        Selections.CreateNewSelection processParameters
        Process_SelectMenu = True
        
    ' This is a dummy entry; it only exists so that Undo/Redo data is correctly generated when a selection is resized
    ElseIf Strings.StringsEqual(processID, "Resize selection", True) Then
        Selections.CreateNewSelection processParameters
        Process_SelectMenu = True
    
    ElseIf Strings.StringsEqual(processID, "Select all", True) Then
        Selections.SelectWholeImage
        Process_SelectMenu = True
        
    End If

End Function
```

---

</SwmSnippet>

<SwmSnippet path="/Modules/Processor.bas" line="252">

---

After returning from Process_SelectMenu, Process checks if the command was handled. If not, it passes the processID to Process_AdjustmentsMenu, so adjustment-related commands get a shot at handling it next. This keeps the routing predictable and modular.

```visual basic
    'Adjustment menu operations have been successfully migrated to XML strings.  (None of their functions raise special return conditions, FYI.)
    If (Not processFound) Then processFound = Process_AdjustmentsMenu(processID, raiseDialog, processParameters, createUndo, relevantTool, recordAction, returnDetails)
    
```

---

</SwmSnippet>

## Adjustment Command Dispatch and Execution

```mermaid
%%{init: {"flowchart": {"defaultRenderer": "elk"}} }%%
flowchart TD
    node1["User selects adjustment/filter from menu"]
    click node1 openCode "Modules/Processor.bas:2030:2128"
    node1 --> node2{"Which adjustment/filter is selected?"}
    click node2 openCode "Modules/Processor.bas:2032:2184"
    node2 -->|"Film negative"| node3["Film Negative Effect and Progress Updates"]
    
    node2 -->|"Invert hue"| node4["Hue Inversion Effect and Progress Updates"]
    
    node2 -->|"Invert RGB"| node5["RGB Inversion Effect and Progress Updates"]
    
    node2 -->|"Shift colors left"| node6["RGB Channel Shift Effect and Progress Updates"]
    
    node2 -->|"Shift colors right"| node7["RGB Channel Shift Effect and Progress Updates"]
    
    node2 -->|"Maximum channel"| node8["Channel Max/Min Pixel Processing"]
    
    node2 -->|"Minimum channel"| node9["Channel Max/Min Pixel Processing"]
    
    node2 -->|"Other recognized adjustment"| node10["Show dialog or apply adjustment directly"]
    click node10 openCode "Modules/Processor.bas:2032:2184"
    node2 -->|"Unrecognized adjustment"| node11["No adjustment applied"]
    click node11 openCode "Modules/Processor.bas:2197:2199"

classDef HeadingStyle fill:#777777,stroke:#333,stroke-width:2px;
click node3 goToHeading "Film Negative Effect and Progress Updates"
node3:::HeadingStyle
click node4 goToHeading "Hue Inversion Effect and Progress Updates"
node4:::HeadingStyle
click node5 goToHeading "RGB Inversion Effect and Progress Updates"
node5:::HeadingStyle
click node6 goToHeading "RGB Channel Shift Effect and Progress Updates"
node6:::HeadingStyle
click node7 goToHeading "RGB Channel Shift Effect and Progress Updates"
node7:::HeadingStyle
click node8 goToHeading "Channel Max/Min Pixel Processing"
node8:::HeadingStyle
click node9 goToHeading "Channel Max/Min Pixel Processing"
node9:::HeadingStyle
```

<SwmSnippet path="/Modules/Processor.bas" line="2030">

---

In Process_AdjustmentsMenu, we match the processID to a known adjustment command. For each match, we either show a dialog (if raiseDialog is set) or call the adjustment function directly with the parameters. This covers everything from brightness/contrast to color balance and more. Interface.bas is called next for dialog-based adjustments, so the user can tweak settings before applying them.

```visual basic
Private Function Process_AdjustmentsMenu(ByVal processID As String, Optional raiseDialog As Boolean = False, Optional processParameters As String = vbNullString, Optional createUndo As PD_UndoType = UNDO_Nothing, Optional relevantTool As Long = -1, Optional recordAction As Boolean = True, Optional ByRef returnDetails As String = vbNullString) As Boolean
    
    If Strings.StringsEqual(processID, "Auto correct", True) Then
        Filters_Adjustments.AutoCorrectImage
        Process_AdjustmentsMenu = True
    
    ElseIf Strings.StringsEqual(processID, "Auto enhance", True) Then
        Filters_Adjustments.fxAutoEnhance
        Process_AdjustmentsMenu = True
    
    'Luminance adjustment functions
    ElseIf Strings.StringsEqual(processID, "Brightness and contrast", True) Then
        If raiseDialog Then ShowPDDialog vbModal, FormBrightnessContrast Else FormBrightnessContrast.BrightnessContrast processParameters
        Process_AdjustmentsMenu = True
    
    ElseIf Strings.StringsEqual(processID, "Curves", True) Then
        If raiseDialog Then ShowPDDialog vbModal, FormCurves Else FormCurves.ApplyCurveToImage processParameters
        Process_AdjustmentsMenu = True
    
    ElseIf Strings.StringsEqual(processID, "Dehaze", True) Then
        If raiseDialog Then ShowPDDialog vbModal, FormDehaze Else FormDehaze.ApplyDehaze processParameters
        Process_AdjustmentsMenu = True
        
    ElseIf Strings.StringsEqual(processID, "Exposure", True) Then
        If raiseDialog Then ShowPDDialog vbModal, FormExposure Else FormExposure.Exposure processParameters
        Process_AdjustmentsMenu = True
        
    ElseIf Strings.StringsEqual(processID, "Gamma", True) Then
        If raiseDialog Then ShowPDDialog vbModal, FormGamma Else FormGamma.GammaCorrect processParameters
        Process_AdjustmentsMenu = True
    
    ElseIf Strings.StringsEqual(processID, "HDR", True) Then
        If raiseDialog Then ShowPDDialog vbModal, FormHDR Else FormHDR.ApplyImitationHDR processParameters
        Process_AdjustmentsMenu = True
    
    ElseIf Strings.StringsEqual(processID, "Levels", True) Then
        If raiseDialog Then ShowPDDialog vbModal, FormLevels Else FormLevels.MapImageLevels processParameters
        Process_AdjustmentsMenu = True
        
    ElseIf Strings.StringsEqual(processID, "Shadows and highlights", True) Or Strings.StringsEqual(processID, "Shadow and highlight", True) Then
        If raiseDialog Then ShowPDDialog vbModal, FormShadowHighlight Else FormShadowHighlight.ApplyShadowHighlight processParameters
        Process_AdjustmentsMenu = True
        
    ElseIf Strings.StringsEqual(processID, "White balance", True) Then
        If raiseDialog Then ShowPDDialog vbModal, FormWhiteBalance Else Filters_Adjustments.AutoWhiteBalance processParameters
        Process_AdjustmentsMenu = True
    
    'Color adjustments
    ElseIf Strings.StringsEqual(processID, "Color balance", True) Then
        If raiseDialog Then ShowPDDialog vbModal, FormColorBalance Else FormColorBalance.ApplyColorBalance processParameters
        Process_AdjustmentsMenu = True
    
    ElseIf Strings.StringsEqual(processID, "Color lookup", True) Then
        If raiseDialog Then ShowPDDialog vbModal, FormColorLookup Else FormColorLookup.ApplyColorLookupEffect processParameters
        Process_AdjustmentsMenu = True
    
    ElseIf Strings.StringsEqual(processID, "Colorize", True) Then
        If raiseDialog Then ShowPDDialog vbModal, FormColorize Else FormColorize.ColorizeImage processParameters
        Process_AdjustmentsMenu = True
    
    ElseIf Strings.StringsEqual(processID, "Hue and saturation", True) Then
        If raiseDialog Then ShowPDDialog vbModal, FormHSL Else FormHSL.AdjustImageHSL processParameters
        Process_AdjustmentsMenu = True
    
    ElseIf Strings.StringsEqual(processID, "Photo filter", True) Then
        If raiseDialog Then ShowPDDialog vbModal, FormPhotoFilters Else FormPhotoFilters.ApplyPhotoFilter processParameters
        Process_AdjustmentsMenu = True
    
    ElseIf Strings.StringsEqual(processID, "Replace color", True) Then
        If raiseDialog Then ShowPDDialog vbModal, FormReplaceColor Else FormReplaceColor.ReplaceSelectedColor processParameters
        Process_AdjustmentsMenu = True
    
    ElseIf Strings.StringsEqual(processID, "Sepia", True) Then
        If raiseDialog Then ShowPDDialog vbModal, FormSepia Else FormSepia.ApplySepiaEffect processParameters
        Process_AdjustmentsMenu = True
        
    ElseIf Strings.StringsEqual(processID, "Split toning", True) Then
        If raiseDialog Then ShowPDDialog vbModal, FormSplitTone Else FormSplitTone.SplitTone processParameters
        Process_AdjustmentsMenu = True
        
    ElseIf Strings.StringsEqual(processID, "Temperature", True) Then
        If raiseDialog Then ShowPDDialog vbModal, FormColorTemp Else FormColorTemp.ApplyTemperatureToImage processParameters
        Process_AdjustmentsMenu = True
        
    ElseIf Strings.StringsEqual(processID, "Tint", True) Then
        If raiseDialog Then ShowPDDialog vbModal, FormTint Else FormTint.AdjustTint processParameters
        Process_AdjustmentsMenu = True
        
    ElseIf Strings.StringsEqual(processID, "Vibrance", True) Then
        If raiseDialog Then ShowPDDialog vbModal, FormVibrance Else FormVibrance.Vibrance processParameters
        Process_AdjustmentsMenu = True
    
    'Grayscale conversions
    ElseIf Strings.StringsEqual(processID, "Black and white", True) Then
        If raiseDialog Then ShowPDDialog vbModal, FormGrayscale Else FormGrayscale.GrayscaleConvert_Central processParameters
        Process_AdjustmentsMenu = True
    
```

---

</SwmSnippet>

<SwmSnippet path="/Modules/Processor.bas" line="2127">

---

In Process_AdjustmentsMenu, when the processID is 'Film negative', we call MenuNegative from Filters_Color.bas to run the negative filter. This hands off the pixel manipulation to a specialized function for that effect.

```visual basic
    'Invert operations
    ElseIf Strings.StringsEqual(processID, "Film negative", True) Then
        MenuNegative
        Process_AdjustmentsMenu = True
    
    ElseIf Strings.StringsEqual(processID, "Invert hue", True) Then
```

---

</SwmSnippet>

### Film Negative Effect and Progress Updates

```mermaid
%%{init: {"flowchart": {"defaultRenderer": "elk"}} }%%
flowchart TD
  node1["Inform user: Calculating film negative values"]
  click node1 openCode "Modules/Filters_Color.bas:169:171"
  node1 --> node2["Prepare image data"]
  click node2 openCode "Modules/Filters_Color.bas:173:176"
  node2 --> node3["Set up progress bar"]
  click node3 openCode "Modules/Filters_Color.bas:187:189"
  node3 --> node4["Start pixel processing"]
  click node4 openCode "Modules/Filters_Color.bas:197:221"
  subgraph loop1["For each pixel in selected area"]
    node4 --> node5["Transform pixel to negative"]
    click node5 openCode "Modules/Filters_Color.bas:200:214"
    node5 --> node6{"Did user press ESC?"}
    click node6 openCode "Modules/Filters_Color.bas:218:218"
    node6 -->|"Yes"| node10["Cancel operation"]
    click node10 openCode "Modules/Filters_Color.bas:218:218"
    node6 -->|"No"| node7["Update progress bar"]
    click node7 openCode "Modules/Filters_Color.bas:219:220"
    node7 --> node4
  end
  node4 --> node8["Finalize image data"]
  click node8 openCode "Modules/Filters_Color.bas:224:227"
  node10 --> node9["End"]
  click node9 openCode "Modules/Filters_Color.bas:229:229"
classDef HeadingStyle fill:#777777,stroke:#333,stroke-width:2px;
```

<SwmSnippet path="/Modules/Filters_Color.bas" line="169">

---

In MenuNegative, we kick off the effect by showing a message to the user so they know the negative filter is running. This uses Interface.bas to display the status.

```visual basic
Public Sub MenuNegative()

    Message "Calculating film negative values..."

```

---

</SwmSnippet>

<SwmSnippet path="/Modules/Interface.bas" line="1668">

---

Message checks for duplicates before displaying anything, translates the string if needed, and appends a 'Recording' tag if macros are active. It then shows the message on the canvas unless we're in batch mode.

```visual basic
Public Sub Message(ByVal mString As String, ParamArray ExtraText() As Variant)

    Dim i As Long

    'Before doing anything else, check for a duplicate message request.  They are automatically ignored.
    Dim tmpDupeCheckString As String
    tmpDupeCheckString = mString
    
    If (UBound(ExtraText) >= LBound(ExtraText)) Then
        
        For i = LBound(ExtraText) To UBound(ExtraText)
            If Strings.StringsNotEqual(CStr(ExtraText(i)), "DONOTLOG", True) Then
                tmpDupeCheckString = Replace$(tmpDupeCheckString, "%" & CStr(i + 1), CStr(ExtraText(i)))
            End If
        Next i
        
    End If
    
    'If the message request is for a novel string (e.g. one that differs from the previous message request), display it.
    ' Otherwise, exit now.
    If Strings.StringsNotEqual(m_PrevMessage, tmpDupeCheckString, False) Then
        
        'In debug mode, mirror the message output to PD's central Debugger.  Note that this behavior can be overridden by
        ' supplying the string "DONOTLOG" as the final entry in the ParamArray.
        If UserPrefs.GenerateDebugLogs Then
        
            If (UBound(ExtraText) < LBound(ExtraText)) Then
                PDDebug.LogAction tmpDupeCheckString, PDM_User_Message
            Else
            
                'Check the last param passed.  If it's the string "DONOTLOG", do not log this entry.  (PD sometimes uses this
                ' to avoid logging useless data, like layer hover events or download updates.)
                If Strings.StringsNotEqual(CStr(ExtraText(UBound(ExtraText))), "DONOTLOG", False) Then
                    PDDebug.LogAction tmpDupeCheckString, PDM_User_Message
                End If
            
            End If
        
        End If
        
        'Cache the contents of the untranslated message, so we can check for duplicates on the next message request
        m_PrevMessage = tmpDupeCheckString
                
        Dim newString As String
        newString = mString
    
        'All messages are translatable, but we don't want to translate them if the translation object isn't ready yet.
        ' This only happens for a few messages when the program is first loaded, and at some point, I will eventually getting
        ' around to removing them entirely.
        If (Not g_Language Is Nothing) Then
            If g_Language.ReadyToTranslate Then
                If g_Language.TranslationActive Then newString = g_Language.TranslateMessage(mString)
            End If
        End If
        
        'Once the message is translated, we can add back in any optional text supplied in the ParamArray
        If (UBound(ExtraText) >= LBound(ExtraText)) Then
            For i = LBound(ExtraText) To UBound(ExtraText)
                newString = Replace$(newString, "%" & i + 1, CStr(ExtraText(i)))
            Next i
        End If
        
        'While macros are active, append a "Recording" message to help orient the user
        If (Macros.GetMacroStatus = MacroSTART) Then newString = newString & " {-" & g_Language.TranslateMessage("Recording") & "-}"
        
        'Post the message to the screen
        If (Macros.GetMacroStatus <> MacroBATCH) Then FormMain.MainCanvas(0).DisplayCanvasMessage newString
        
        'Update the global "previous message" string, so external functions can access it.
        m_LastFullMessage = newString
        
    End If
    
End Sub
```

---

</SwmSnippet>

<SwmSnippet path="/Modules/Filters_Color.bas" line="173">

---

SetProgBarVal updates the progress bar and Windows taskbar (if supported), but skips all that if we're running a batch macro. It also processes paint messages to keep the UI responsive during long operations.

```visual basic
    'Create a local array and point it at the pixel data we want to operate on
    Dim imageData() As Byte, tmpSA As SafeArray2D
    EffectPrep.PrepImageData tmpSA
    workingDIB.WrapArrayAroundDIB imageData, tmpSA
    
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = curDIBValues.Left
    initY = curDIBValues.Top
    finalX = curDIBValues.Right
    finalY = curDIBValues.Bottom
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    ProgressBars.SetProgBarMax finalY
    progBarCheck = ProgressBars.FindBestProgBarValue()
    
    'Finally, a bunch of variables used in color calculation
    Dim r As Long, g As Long, b As Long
    Dim h As Double, s As Double, v As Double
    
    'Apply the filter
    initX = initX * 4
    finalX = finalX * 4
    For y = initY To finalY
    For x = initX To finalX Step 4
        
        'Get red, green, and blue values from the array
        b = imageData(x, y)
        g = imageData(x + 1, y)
        r = imageData(x + 2, y)
        
        'Use those to calculate hue and saturation
        Colors.ImpreciseRGBtoHSL r, g, b, h, s, v
        
        'Convert those HSL values back to RGB, but substitute inverted luminance
        Colors.ImpreciseHSLtoRGB h, s, 1# - v, r, g, b
        
        'Assign the new RGB values back into the array
        imageData(x, y) = b
        imageData(x + 1, y) = g
        imageData(x + 2, y) = r
        
    Next x
        If (y And progBarCheck) = 0 Then
            If Interface.UserPressedESC() Then Exit For
            SetProgBarVal y
        End If
    Next y
        
```

---

</SwmSnippet>

<SwmSnippet path="/Modules/ProgressBars.bas" line="43">

---

SetProgBarVal updates the progress bar and Windows taskbar (if supported), but skips all that if we're running a batch macro. It also processes paint messages to keep the UI responsive during long operations.

```visual basic
Public Sub SetProgBarVal(ByVal pbVal As Double)
    
    If (Macros.GetMacroStatus <> MacroBATCH) Then
        
        FormMain.MainCanvas(0).ProgBar_SetValue pbVal
        
        'On Windows 7 (or later), we also update the taskbar to reflect the current progress
        If OS.IsWin7OrLater Then OS.SetTaskbarProgressValue pbVal, GetProgBarMax, FormMain.hWnd
        
        'Process some window messages on the main form, to prevent the dreaded "Not Responding" state
        ' when PD is in the midst of a long-running action.
        VBHacks.DoEvents_PaintOnly False
        
    End If
    
End Sub
```

---

</SwmSnippet>

<SwmSnippet path="/Modules/Filters_Color.bas" line="223">

---

After updating progress, MenuNegative deallocates the imageData array and calls EffectPrep.FinalizeImageData to finish up rendering and cleanup.

```visual basic
    'Safely deallocate imageData()
    workingDIB.UnwrapArrayFromDIB imageData
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering
    EffectPrep.FinalizeImageData
    
End Sub
```

---

</SwmSnippet>

### Hue and RGB Inversion Routing

```mermaid
%%{init: {"flowchart": {"defaultRenderer": "elk"}} }%%
flowchart TD
    node1["User selects adjustment from Adjustments menu"] --> node2{"processID: Which adjustment?"}
    click node1 openCode "Modules/Processor.bas:2133:2136"
    click node2 openCode "Modules/Processor.bas:2133:2136"
    node2 -->|"Invert Hue"| node3["Apply hue inversion to image"]
    click node3 openCode "Modules/Processor.bas:2133:2134"
    node2 -->|"Invert RGB"| node4["Apply RGB inversion to image"]
    click node4 openCode "Modules/Processor.bas:2136:2136"
    node2 -->|"Other/Unsupported"| node5["No adjustment applied"]
    click node5 openCode "Modules/Processor.bas:2133:2136"
    node3 --> node6["Adjustment applied"]
    node4 --> node6
    node5 --> node6
    click node6 openCode "Modules/Processor.bas:2134:2136"
classDef HeadingStyle fill:#777777,stroke:#333,stroke-width:2px;
```

<SwmSnippet path="/Modules/Processor.bas" line="2133">

---

After running MenuNegative, Process_AdjustmentsMenu checks for 'Invert hue' and calls MenuInvertHue from Filters_Color.bas to handle hue inversion. This hands off the pixel work to a specialized function for that effect.

```visual basic
        MenuInvertHue
        Process_AdjustmentsMenu = True
        
    ElseIf Strings.StringsEqual(processID, "Invert RGB", True) Then
```

---

</SwmSnippet>

### Hue Inversion Effect and Progress Updates

```mermaid
%%{init: {"flowchart": {"defaultRenderer": "elk"}} }%%
flowchart TD
    node1["Start hue inversion filter"] --> node2["Select area to process"]
    click node1 openCode "Modules/Filters_Color.bas:232:234"
    click node2 openCode "Modules/Filters_Color.bas:236:246"
    node2 --> node3["Begin hue inversion"]
    click node3 openCode "Modules/Filters_Color.bas:257:259"
    subgraph loop1["For each pixel in selected area"]
        node3 --> node4["Invert pixel hue"]
        click node4 openCode "Modules/Filters_Color.bas:263:273"
        node4 --> node5{"User pressed ESC?"}
        click node5 openCode "Modules/Filters_Color.bas:285:285"
        node5 -->|"Yes"| node6["Cancel operation"]
        click node6 openCode "Modules/Filters_Color.bas:285:285"
        node5 -->|"No"| node4
    end
    node4 --> node7["Finalize and display result"]
    click node7 openCode "Modules/Filters_Color.bas:290:294"
    node7 --> node8["End filter"]
    click node8 openCode "Modules/Filters_Color.bas:296:296"

classDef HeadingStyle fill:#777777,stroke:#333,stroke-width:2px;
```

<SwmSnippet path="/Modules/Filters_Color.bas" line="232">

---

In MenuInvertHue, we start by showing a message to let the user know the hue inversion is happening. This uses Interface.bas for status feedback.

```visual basic
Public Sub MenuInvertHue()

    Message "Inverting..."

```

---

</SwmSnippet>

<SwmSnippet path="/Modules/Filters_Color.bas" line="236">

---

After showing the message, MenuInvertHue wraps the image data, loops through each pixel, converts RGB to HSL, inverts the hue, converts back to RGB, and writes the result. Progress bar updates are optimized to avoid slowing things down.

```visual basic
    'Create a local array and point it at the pixel data we want to operate on
    Dim imageData() As Byte, tmpSA As SafeArray2D
    EffectPrep.PrepImageData tmpSA
    workingDIB.WrapArrayAroundDIB imageData, tmpSA
    
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = curDIBValues.Left
    initY = curDIBValues.Top
    finalX = curDIBValues.Right
    finalY = curDIBValues.Bottom
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    ProgressBars.SetProgBarMax finalY
    progBarCheck = ProgressBars.FindBestProgBarValue()
    
    'Finally, a bunch of variables used in color calculation
    Dim r As Long, g As Long, b As Long
    Dim h As Double, s As Double, l As Double
    
    'Apply the filter
    initX = initX * 4
    finalX = finalX * 4
    For y = initY To finalY
    For x = initX To finalX Step 4
        
        'Get red, green, and blue values from the array
        b = imageData(x, y)
        g = imageData(x + 1, y)
        r = imageData(x + 2, y)
        
        'Use a fast but somewhat imprecise conversion to HSL.  (Note that this returns hue on the
        ' weird range [-1, 5], which allows for performance optimizations but is not intuitive.)
        Colors.ImpreciseRGBtoHSL r, g, b, h, s, l
        
        'Invert hue
        h = 4# - h
        
        'Convert the newly calculated HSL values back to RGB
        Colors.ImpreciseHSLtoRGB h, s, l, r, g, b
        
        'Assign the new RGB values back into the array
        imageData(x, y) = b
        imageData(x + 1, y) = g
        imageData(x + 2, y) = r
        
    Next x
        If (y And progBarCheck) = 0 Then
            If Interface.UserPressedESC() Then Exit For
            SetProgBarVal y
        End If
    Next y
    
```

---

</SwmSnippet>

<SwmSnippet path="/Modules/Filters_Color.bas" line="290">

---

After updating progress, MenuInvertHue deallocates the imageData array and calls EffectPrep.FinalizeImageData to finish up rendering and cleanup.

```visual basic
    'Safely deallocate imageData()
    workingDIB.UnwrapArrayFromDIB imageData
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering
    EffectPrep.FinalizeImageData
    
End Sub
```

---

</SwmSnippet>

### RGB Inversion Routing

<SwmSnippet path="/Modules/Processor.bas" line="2137">

---

After running MenuInvertHue, Process_AdjustmentsMenu checks for 'Invert RGB' and calls MenuInvert from Filters_Color.bas to handle the RGB inversion. This hands off the pixel work to a specialized function for that effect.

```visual basic
        MenuInvert
        Process_AdjustmentsMenu = True
    
```

---

</SwmSnippet>

### RGB Inversion Effect and Progress Updates

```mermaid
%%{init: {"flowchart": {"defaultRenderer": "elk"}} }%%
flowchart TD
    node1["Start color inversion"] --> node2["Prepare selected area"]
    click node1 openCode "Modules/Filters_Color.bas:57:59"
    click node2 openCode "Modules/Filters_Color.bas:61:75"
    node2 --> node3["Process selected area"]
    click node3 openCode "Modules/Filters_Color.bas:77:87"
    
    subgraph loop1["For each row and pixel in selected area (initX, initY, finalX, finalY)"]
        node3 --> node4["Invert pixel colors"]
        click node4 openCode "Modules/Filters_Color.bas:90:93"
        node4 --> node5{"User pressed ESC?"}
        click node5 openCode "Modules/Filters_Color.bas:95:95"
        node5 -->|"No"| node3
        node5 -->|"Yes"| node6["Stop inversion"]
        click node6 openCode "Modules/Filters_Color.bas:95:95"
    end
    node3 --> node7["Finalize image"]
    click node7 openCode "Modules/Filters_Color.bas:104:105"

classDef HeadingStyle fill:#777777,stroke:#333,stroke-width:2px;
```

<SwmSnippet path="/Modules/Filters_Color.bas" line="57">

---

In MenuInvert, we start by showing a message to let the user know the RGB inversion is happening. This uses Interface.bas for status feedback.

```visual basic
Public Sub MenuInvert()
        
    Message "Inverting..."
    
```

---

</SwmSnippet>

<SwmSnippet path="/Modules/Filters_Color.bas" line="61">

---

After showing the message, MenuInvert wraps the image data, loops through each pixel, and inverts the RGB channels using XOR. The alpha channel is untouched. Progress is updated as we go, and the pixel array is deallocated and finalized at the end.

```visual basic
    'Create a local array and point it at the pixel data we want to operate on
    Dim imageData() As Byte, tmpSA As SafeArray2D
    EffectPrep.PrepImageData tmpSA
    
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = curDIBValues.Left
    initY = curDIBValues.Top
    finalX = curDIBValues.Right
    finalY = curDIBValues.Bottom
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    ProgressBars.SetProgBarMax finalY
    progBarCheck = ProgressBars.FindBestProgBarValue()
    
    Dim tmpSA1D As SafeArray1D, pxData As Long, pxStride As Long
    workingDIB.WrapArrayAroundScanline imageData, tmpSA1D, initY
    pxData = tmpSA1D.pvData
    pxStride = tmpSA1D.cElements
    
    'Images are always 32-bpp
    initX = initX * 4
    finalX = finalX * 4
    
    'After all that work, the Invert code itself is very small and unexciting!
    For y = initY To finalY
        tmpSA1D.pvData = pxData + pxStride * y
    For x = initX To finalX Step 4
        imageData(x) = 255 Xor imageData(x)
        imageData(x + 1) = 255 Xor imageData(x + 1)
        imageData(x + 2) = 255 Xor imageData(x + 2)
    Next x
        If (y And progBarCheck) = 0 Then
            If Interface.UserPressedESC() Then Exit For
            ProgressBars.SetProgBarVal y
        End If
    Next y
    
    'Safely deallocate imageData()
    workingDIB.UnwrapArrayFromDIB imageData
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering
    EffectPrep.FinalizeImageData
    
End Sub
```

---

</SwmSnippet>

### Map, Monochrome, and Channel Adjustment Routing

```mermaid
%%{init: {"flowchart": {"defaultRenderer": "elk"}} }%%
flowchart TD
    node1["User selects adjustment operation"] --> node2{"Which adjustment?"}
    click node1 openCode "Modules/Processor.bas:2140:2180"
    node2 -->|"Gradient map"| node3{"Show dialog?"}
    node2 -->|"Palette map"| node4{"Show dialog?"}
    node2 -->|"Color to monochrome"| node5{"Show dialog?"}
    node2 -->|"Monochrome to gray"| node6{"Show dialog?"}
    node2 -->|"Channel mixer"| node7{"Show dialog?"}
    node2 -->|"Rechannel"| node8{"Show dialog?"}
    node2 -->|"Shift colors"| node9{"Left or Right?"}
    node2 -->|"Max/Min channel"| node10{"Max or Min?"}
    click node2 openCode "Modules/Processor.bas:2140:2180"
    node3 -->|"Yes"| node11["Show dialog for Gradient map"]
    node3 -->|"No"| node12["Apply Gradient map effect"]
    click node3 openCode "Modules/Processor.bas:2141:2143"
    node4 -->|"Yes"| node13["Show dialog for Palette map"]
    node4 -->|"No"| node14["Apply Palette map effect"]
    click node4 openCode "Modules/Processor.bas:2145:2147"
    node5 -->|"Yes"| node15["Show dialog for Monochrome"]
    node5 -->|"No"| node16["Apply Monochrome effect"]
    click node5 openCode "Modules/Processor.bas:2151:2153"
    node6 -->|"Yes"| node17["Show dialog for Mono to gray"]
    node6 -->|"No"| node18["Apply Mono to gray effect"]
    click node6 openCode "Modules/Processor.bas:2155:2157"
    node7 -->|"Yes"| node19["Show dialog for Channel mixer"]
    node7 -->|"No"| node20["Apply Channel mixer effect"]
    click node7 openCode "Modules/Processor.bas:2160:2162"
    node8 -->|"Yes"| node21["Show dialog for Rechannel"]
    node8 -->|"No"| node22["Apply Rechannel effect"]
    click node8 openCode "Modules/Processor.bas:2164:2166"
    node9 -->|"Left"| node23["Shift colors left"]
    node9 -->|"Right"| node24["Shift colors right"]
    click node9 openCode "Modules/Processor.bas:2168:2174"
    node10 -->|"Max"| node25["Apply Maximum channel"]
    node10 -->|"Min"| node26["Apply Minimum channel"]
    click node10 openCode "Modules/Processor.bas:2176:2180"

classDef HeadingStyle fill:#777777,stroke:#333,stroke-width:2px;
```

<SwmSnippet path="/Modules/Processor.bas" line="2140">

---

After returning from Filters_Color.bas, Process_AdjustmentsMenu checks if the next map, monochrome, or channel adjustment needs a dialog. If so, it shows the right form from Interface.bas; otherwise, it just runs the effect. This lets users interact when needed, or automates the process if not.

```visual basic
    'Map operations
    ElseIf Strings.StringsEqual(processID, "Gradient map", True) Then
        If raiseDialog Then ShowPDDialog vbModal, FormGradientMap Else FormGradientMap.ApplyGradientMap processParameters
        Process_AdjustmentsMenu = True
        
    ElseIf Strings.StringsEqual(processID, "Palette map", True) Then
        If raiseDialog Then ShowPDDialog vbModal, FormPalettize Else FormPalettize.ApplyPalettizeEffect processParameters
        Process_AdjustmentsMenu = True
        
    'Monochrome conversion
    ' (Note: all monochrome conversion operations are condensed into a single function.  (Past versions spread them across multiple functions.))
    ElseIf Strings.StringsEqual(processID, "Color to monochrome", True) Then
        If raiseDialog Then ShowPDDialog vbModal, FormMonochrome Else FormMonochrome.MonochromeConvert_Central processParameters
        Process_AdjustmentsMenu = True
        
    ElseIf Strings.StringsEqual(processID, "Monochrome to gray", True) Then
        If raiseDialog Then ShowPDDialog vbModal, FormMonoToColor Else FormMonoToColor.ConvertMonoToColor processParameters
        Process_AdjustmentsMenu = True
        
    'Channel operations
    ElseIf Strings.StringsEqual(processID, "Channel mixer", True) Then
        If raiseDialog Then ShowPDDialog vbModal, FormChannelMixer Else FormChannelMixer.ApplyChannelMixer processParameters
        Process_AdjustmentsMenu = True
        
    ElseIf Strings.StringsEqual(processID, "Rechannel", True) Then
        If raiseDialog Then ShowPDDialog vbModal, FormRechannel Else FormRechannel.RechannelImage processParameters
        Process_AdjustmentsMenu = True
        
    ElseIf Strings.StringsEqual(processID, "Shift colors (left)", True) Then
        MenuCShift True
        Process_AdjustmentsMenu = True
        
    ElseIf Strings.StringsEqual(processID, "Shift colors (right)", True) Then
```

---

</SwmSnippet>

<SwmSnippet path="/Modules/Processor.bas" line="2173">

---

After returning from Interface.bas, Process_AdjustmentsMenu checks for color shift and channel max/min commands, and calls the right function in Filters_Color.bas to handle the pixel manipulation.

```visual basic
        MenuCShift False
        Process_AdjustmentsMenu = True
                
    ElseIf Strings.StringsEqual(processID, "Maximum channel", True) Then
        FilterMaxMinChannel True
        Process_AdjustmentsMenu = True
    
    ElseIf Strings.StringsEqual(processID, "Minimum channel", True) Then
```

---

</SwmSnippet>

### RGB Channel Shift Effect and Progress Updates

```mermaid
%%{init: {"flowchart": {"defaultRenderer": "elk"}} }%%
flowchart TD
    node1["User chooses to shift image colors"] --> node2["Prepare image for editing"]
    click node1 openCode "Modules/Filters_Color.bas:109:111"
    click node2 openCode "Modules/Filters_Color.bas:113:116"
    
    subgraph loop1["For each pixel in selected area"]
        node3{"Shift direction (shiftLeft)?"}
        click node3 openCode "Modules/Filters_Color.bas:139:147"
        node3 -->|"Left"| node4["Shift colors left"]
        click node4 openCode "Modules/Filters_Color.bas:140:142"
        node3 -->|"Right"| node5["Shift colors right"]
        click node5 openCode "Modules/Filters_Color.bas:144:146"
        node4 --> node6["Check for user cancel/progress update"]
        node5 --> node6
        node6{"Continue or cancel?"}
        click node6 openCode "Modules/Filters_Color.bas:154:157"
        node6 -->|"Continue"| node3
        node6 -->|"Cancel"| node8["Stop processing"]
        click node8 openCode "Modules/Filters_Color.bas:155:155"
    end
    loop1 --> node7["Apply changes to image"]
    click node7 openCode "Modules/Filters_Color.bas:160:164"

classDef HeadingStyle fill:#777777,stroke:#333,stroke-width:2px;
```

<SwmSnippet path="/Modules/Filters_Color.bas" line="109">

---

In MenuCShift, we start by showing a message to let the user know the RGB channel shift is happening. This uses Interface.bas for status feedback.

```visual basic
Public Sub MenuCShift(Optional ByVal shiftLeft As Boolean = False)
    
    Message "Shifting RGB values..."
    
```

---

</SwmSnippet>

<SwmSnippet path="/Modules/Filters_Color.bas" line="113">

---

After showing the message, MenuCShift wraps the image data, loops through each pixel, and shifts the RGB channels left or right depending on the flag. Progress bar updates are optimized, and the pixel array is deallocated and finalized at the end.

```visual basic
    'Create a local array and point it at the pixel data we want to operate on
    Dim imageData() As Byte, tmpSA As SafeArray2D
    EffectPrep.PrepImageData tmpSA
    workingDIB.WrapArrayAroundDIB imageData, tmpSA
    
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = curDIBValues.Left
    initY = curDIBValues.Top
    finalX = curDIBValues.Right
    finalY = curDIBValues.Bottom
    
    Dim xStride As Long
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    progBarCheck = ProgressBars.FindBestProgBarValue()
    
    'Finally, a bunch of variables used in color calculation
    Dim r As Long, g As Long, b As Long
    
    'After all that work, the Invert code itself is very small and unexciting!
    For x = initX To finalX
        xStride = x * 4
    For y = initY To finalY
        
        If shiftLeft Then
            g = imageData(xStride, y)
            r = imageData(xStride + 1, y)
            b = imageData(xStride + 2, y)
        Else
            r = imageData(xStride, y)
            b = imageData(xStride + 1, y)
            g = imageData(xStride + 2, y)
        End If
        
        imageData(xStride, y) = b
        imageData(xStride + 1, y) = g
        imageData(xStride + 2, y) = r
        
    Next y
        If (x And progBarCheck) = 0 Then
            If Interface.UserPressedESC() Then Exit For
            SetProgBarVal x
        End If
    Next x
        
```

---

</SwmSnippet>

<SwmSnippet path="/Modules/Filters_Color.bas" line="160">

---

After updating progress, MenuCShift deallocates the imageData array and calls EffectPrep.FinalizeImageData to finish up rendering and cleanup.

```visual basic
    'Safely deallocate imageData()
    workingDIB.UnwrapArrayFromDIB imageData
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering
    EffectPrep.FinalizeImageData
    
End Sub
```

---

</SwmSnippet>

### Channel Max/Min Adjustment Routing

<SwmSnippet path="/Modules/Processor.bas" line="2181">

---

After MenuCShift, Process_AdjustmentsMenu checks for channel max/min adjustments and calls FilterMaxMinChannel if needed, then marks the adjustment as handled.

```visual basic
        FilterMaxMinChannel False
        Process_AdjustmentsMenu = True
        
```

---

</SwmSnippet>

### Channel Max/Min Pixel Processing

```mermaid
%%{init: {"flowchart": {"defaultRenderer": "elk"}} }%%
flowchart TD
    node1["User selects to isolate maximum or minimum color channel"]
    click node1 openCode "Modules/Filters_Color.bas:299:305"
    subgraph loop1["For each pixel in selected area"]
      node2{"Isolate maximum or minimum channel?"}
      click node2 openCode "Modules/Filters_Color.bas:335:345"
      node2 -->|"Maximum"| node3["Keep only the dominant color channel"]
      click node3 openCode "Modules/Filters_Color.bas:336:340"
      node2 -->|"Minimum"| node4["Keep only the least dominant color channel"]
      click node4 openCode "Modules/Filters_Color.bas:341:344"
    end
    loop1 --> node5["Finalize and update image"]
    click node5 openCode "Modules/Filters_Color.bas:358:363"
classDef HeadingStyle fill:#777777,stroke:#333,stroke-width:2px;
```

<SwmSnippet path="/Modules/Filters_Color.bas" line="299">

---

FilterMaxMinChannel starts by telling the user what it's about to do (max or min channel isolation) before jumping into pixel processing.

```visual basic
Public Sub FilterMaxMinChannel(ByVal useMax As Boolean)
    
    If useMax Then
        Message "Isolating maximum color channels..."
    Else
        Message "Isolating minimum color channels..."
    End If
    
```

---

</SwmSnippet>

<SwmSnippet path="/Modules/Filters_Color.bas" line="307">

---

After the message, FilterMaxMinChannel loops through each pixel, using Max3Int/Min3Int to isolate the strongest or weakest channel, and only updates the progress bar occasionally to keep things fast.

```visual basic
    'Create a local array and point it at the pixel data we want to operate on
    Dim imageData() As Byte, tmpSA As SafeArray2D
    EffectPrep.PrepImageData tmpSA
    workingDIB.WrapArrayAroundDIB imageData, tmpSA
    
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = curDIBValues.Left * 4
    initY = curDIBValues.Top
    finalX = curDIBValues.Right * 4
    finalY = curDIBValues.Bottom
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    ProgressBars.SetProgBarMax finalY
    progBarCheck = ProgressBars.FindBestProgBarValue()
    
    'Finally, a bunch of variables used in color calculation
    Dim r As Long, g As Long, b As Long, maxVal As Long, minVal As Long
        
    'Apply the filter
    For y = initY To finalY
    For x = initX To finalX Step 4
        
        b = imageData(x, y)
        g = imageData(x + 1, y)
        r = imageData(x + 2, y)
        
        If useMax Then
            maxVal = Max3Int(r, g, b)
            If r < maxVal Then r = 0
            If g < maxVal Then g = 0
            If b < maxVal Then b = 0
        Else
```

---

</SwmSnippet>

<SwmSnippet path="/Modules/PDMath.bas" line="607">

---

Max3Int just picks the largest of three values using nested Ifs. It's a helper for the channel isolation filter, since VB6 doesn't have a built-in max-of-three.

```visual basic
Public Function Max3Int(ByVal rR As Long, ByVal rG As Long, ByVal rB As Long) As Long
    If (rR > rG) Then
        If (rR > rB) Then Max3Int = rR Else Max3Int = rB
    Else
        If (rB > rG) Then Max3Int = rB Else Max3Int = rG
    End If
End Function
```

---

</SwmSnippet>

<SwmSnippet path="/Modules/Filters_Color.bas" line="341">

---

Back in FilterMaxMinChannel, after handling the max branch, we use Min3Int (from PDMath.bas) to find the weakest channel for the min branch, zeroing out the others. This keeps the logic symmetric and lets the filter handle both modes.

```visual basic
            minVal = Min3Int(r, g, b)
            If r > minVal Then r = 0
            If g > minVal Then g = 0
            If b > minVal Then b = 0
        End If
        
        imageData(x, y) = b
        imageData(x + 1, y) = g
        imageData(x + 2, y) = r
        
    Next x
        If (y And progBarCheck) = 0 Then
            If Interface.UserPressedESC() Then Exit For
```

---

</SwmSnippet>

<SwmSnippet path="/Modules/PDMath.bas" line="633">

---

Min3Int is just like Max3Int, but it finds the smallest of three numbers. It's used by the filter to isolate the weakest channel per pixel.

```visual basic
Public Function Min3Int(ByVal rR As Long, ByVal rG As Long, ByVal rB As Long) As Long
    If (rR < rG) Then
        If (rR < rB) Then Min3Int = rR Else Min3Int = rB
    Else
        If (rB < rG) Then Min3Int = rB Else Min3Int = rG
    End If
End Function
```

---

</SwmSnippet>

<SwmSnippet path="/Modules/Filters_Color.bas" line="354">

---

After the pixel math, FilterMaxMinChannel only updates the progress bar every so often (using progBarCheck) and checks for ESC to let the user cancel. This keeps things responsive without bogging down the filter.

```visual basic
            SetProgBarVal y
        End If
    Next y
        
```

---

</SwmSnippet>

<SwmSnippet path="/Modules/Filters_Color.bas" line="358">

---

After the pixel loop, FilterMaxMinChannel unwraps the image data array and calls EffectPrep.FinalizeImageData to finish up rendering and cleanup. This makes the changes visible and wraps up the filter.

```visual basic
    'Safely deallocate imageData()
    workingDIB.UnwrapArrayFromDIB imageData
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering
    EffectPrep.FinalizeImageData
    
End Sub
```

---

</SwmSnippet>

### Histogram and Equalization Adjustment Routing

```mermaid
%%{init: {"flowchart": {"defaultRenderer": "elk"}} }%%
flowchart TD
    node1{"processID: Which histogram operation did the user select?"}
    click node1 openCode "Modules/Processor.bas:2184:2197"
    node1 -->|"Display histogram"| node2["Show histogram dialog"]
    click node2 openCode "Modules/Processor.bas:2186:2187"
    node1 -->|"Stretch histogram"| node3["Stretch histogram"]
    click node3 openCode "Modules/Processor.bas:2190:2191"
    node1 -->|"Equalize"| node4{"raiseDialog: Show dialog or run equalize directly?"}
    click node4 openCode "Modules/Processor.bas:2194:2195"
    node4 -->|"Show dialog"| node5["Show equalize dialog"]
    click node5 openCode "Modules/Processor.bas:2194:2195"
    node4 -->|"Run directly"| node6["Equalize histogram"]
    click node6 openCode "Modules/Processor.bas:2194:2195"
    node2 --> node7["Operation complete, return to menu"]
    node3 --> node7
    node5 --> node7
    node6 --> node7
    click node7 openCode "Modules/Processor.bas:2187:2196"

classDef HeadingStyle fill:#777777,stroke:#333,stroke-width:2px;
```

<SwmSnippet path="/Modules/Processor.bas" line="2184">

---

After returning from Filters_Color.bas, Process_AdjustmentsMenu wraps up by routing histogram and equalization adjustments. It shows dialogs for interactive commands or runs the effect directly, then marks the adjustment as handled. This keeps the adjustment flow modular and supports both manual and automated workflows.

```visual basic
    'Histogram functions
    ElseIf Strings.StringsEqual(processID, "Display histogram", True) Then
        ShowPDDialog vbModal, FormHistogram
        Process_AdjustmentsMenu = True
    
    ElseIf Strings.StringsEqual(processID, "Stretch histogram", True) Then
        Histograms.StretchHistogram
        Process_AdjustmentsMenu = True
        
    ElseIf Strings.StringsEqual(processID, "Equalize", True) Then
        If raiseDialog Then ShowPDDialog vbModal, FormEqualize Else FormEqualize.EqualizeHistogram processParameters
        Process_AdjustmentsMenu = True
        
    End If
    
End Function
```

---

</SwmSnippet>

## Effects Menu Command Routing

<SwmSnippet path="/Modules/Processor.bas" line="255">

---

If adjustments didn't handle the command, Process tries the effects menu next.

```visual basic
    'Effects menu operations have been successfully migrated to XML strings.  (None of their functions raise special return conditions, FYI.)
    If (Not processFound) Then processFound = Process_EffectsMenu(processID, raiseDialog, processParameters, createUndo, relevantTool, recordAction, returnDetails)
    
```

---

</SwmSnippet>

## Effect Dispatch and Dialog Routing

<SwmSnippet path="/Modules/Processor.bas" line="1640">

---

In Process_EffectsMenu, we match the processID to a known effect. If raiseDialog is set, we show a dialog for user input; otherwise, we run the effect directly. This pattern repeats for all supported effects, keeping things flexible for both manual and automated workflows.

```visual basic
Private Function Process_EffectsMenu(ByVal processID As String, Optional raiseDialog As Boolean = False, Optional processParameters As String = vbNullString, Optional createUndo As PD_UndoType = UNDO_Nothing, Optional relevantTool As Long = -1, Optional recordAction As Boolean = True, Optional ByRef returnDetails As String = vbNullString) As Boolean

    'Artistic
    If Strings.StringsEqual(processID, "Colored pencil", True) Then
        If raiseDialog Then ShowPDDialog vbModal, FormPencil Else FormPencil.fxColoredPencil processParameters
        Process_EffectsMenu = True
        
    ElseIf Strings.StringsEqual(processID, "Comic book", True) Then
        If raiseDialog Then ShowPDDialog vbModal, FormComicBook Else FormComicBook.fxComicBook processParameters
        Process_EffectsMenu = True
        
    ElseIf Strings.StringsEqual(processID, "Figured glass", True) Then
        If raiseDialog Then ShowPDDialog vbModal, FormFiguredGlass Else FormFiguredGlass.FiguredGlassFX processParameters
        Process_EffectsMenu = True
        
    ElseIf Strings.StringsEqual(processID, "Film noir", True) Then
        If raiseDialog Then ShowPDDialog vbModal, FormFilmNoir Else FormFilmNoir.fxFilmNoir processParameters
        Process_EffectsMenu = True
        
    ElseIf Strings.StringsEqual(processID, "Glass tiles", True) Then
        If raiseDialog Then ShowPDDialog vbModal, FormGlassTiles Else FormGlassTiles.GlassTiles processParameters
        Process_EffectsMenu = True
        
    ElseIf Strings.StringsEqual(processID, "Kaleidoscope", True) Then
        If raiseDialog Then ShowPDDialog vbModal, FormKaleidoscope Else FormKaleidoscope.KaleidoscopeImage processParameters
        Process_EffectsMenu = True
        
    ElseIf Strings.StringsEqual(processID, "Modern art", True) Then
        If raiseDialog Then ShowPDDialog vbModal, FormModernArt Else FormModernArt.ApplyModernArt processParameters
        Process_EffectsMenu = True
        
    ElseIf Strings.StringsEqual(processID, "Oil painting", True) Then
        If raiseDialog Then ShowPDDialog vbModal, FormOilPainting Else FormOilPainting.ApplyOilPaintingEffect processParameters
        Process_EffectsMenu = True
        
    ElseIf Strings.StringsEqual(processID, "Plastic wrap", True) Then
        If raiseDialog Then ShowPDDialog vbModal, FormPlasticWrap Else FormPlasticWrap.ApplyPlasticWrap processParameters
        Process_EffectsMenu = True
        
    ElseIf Strings.StringsEqual(processID, "Posterize", True) Then
        If raiseDialog Then ShowPDDialog vbModal, FormPosterize Else FormPosterize.fxPosterize processParameters
        Process_EffectsMenu = True
        
    ElseIf Strings.StringsEqual(processID, "Relief", True) Then
        If raiseDialog Then ShowPDDialog vbModal, FormRelief Else FormRelief.ApplyReliefEffect processParameters
        Process_EffectsMenu = True
        
    ElseIf Strings.StringsEqual(processID, "Stained glass", True) Then
        If raiseDialog Then ShowPDDialog vbModal, FormStainedGlass Else FormStainedGlass.fxStainedGlass processParameters
        Process_EffectsMenu = True
        
    'Blur
    ElseIf Strings.StringsEqual(processID, "Box blur", True) Then
        If raiseDialog Then ShowPDDialog vbModal, FormBoxBlur Else FormBoxBlur.BoxBlurFilter processParameters
        Process_EffectsMenu = True
        
    ElseIf Strings.StringsEqual(processID, "Gaussian blur", True) Then
        If raiseDialog Then ShowPDDialog vbModal, FormGaussianBlur Else FormGaussianBlur.GaussianBlurFilter processParameters
        Process_EffectsMenu = True
        
    ElseIf Strings.StringsEqual(processID, "Surface blur", True) Then
        If raiseDialog Then ShowPDDialog vbModal, FormSurfaceBlur Else FormSurfaceBlur.BilateralFilter_Central processParameters
        Process_EffectsMenu = True
        
    ElseIf Strings.StringsEqual(processID, "Motion blur", True) Then
        If raiseDialog Then ShowPDDialog vbModal, FormMotionBlur Else FormMotionBlur.MotionBlurFilter processParameters
        Process_EffectsMenu = True
        
    ElseIf Strings.StringsEqual(processID, "Radial blur", True) Then
        If raiseDialog Then ShowPDDialog vbModal, FormRadialBlur Else FormRadialBlur.RadialBlurFilter processParameters
        Process_EffectsMenu = True
        
    ElseIf Strings.StringsEqual(processID, "Zoom blur", True) Then
        If raiseDialog Then ShowPDDialog vbModal, FormZoomBlur Else FormZoomBlur.ApplyZoomBlur processParameters
        Process_EffectsMenu = True
        
```

---

</SwmSnippet>

<SwmSnippet path="/Modules/Processor.bas" line="1716">

---

After handling standard effects, Process_EffectsMenu checks for Grid blur. If matched, it calls FilterGridBlur directly (no dialog), since this effect isn't exposed to users and is slated for future migration.

```visual basic
    'TODO: Grid blur (and previously, chroma blur) lived here.  Chroma blur has been removed because I'm going to
    ' re-implement it in LAB space.  Grid blur still exists but is not accessible to the user anywhere; it is going
    ' to be moved to an eventual texture menu, and I don't want to forget about it.
    ElseIf Strings.StringsEqual(processID, "Grid blur", True) Then
        FilterGridBlur
        Process_EffectsMenu = True
    
```

---

</SwmSnippet>

### Grid Blur Pixel Processing

```mermaid
%%{init: {"flowchart": {"defaultRenderer": "elk"}} }%%
flowchart TD
    node1["Start grid blur filter"]
    click node1 openCode "Modules/Filters_Area.bas:277:278"
    subgraph loop1["For each column in image"]
        node2["Calculate vertical color averages"]
        click node2 openCode "Modules/Filters_Area.bas:314:327"
    end
    node1 --> loop1
    subgraph loop2["For each row in image"]
        node3["Calculate horizontal color averages"]
        click node3 openCode "Modules/Filters_Area.bas:330:343"
    end
    loop1 --> loop2
    subgraph loop3["For each pixel in image"]
        node4["Average vertical and horizontal color values"]
        click node4 openCode "Modules/Filters_Area.bas:348:355"
        node4 --> node5{"Is any color value > 255?"}
        click node5 openCode "Modules/Filters_Area.bas:358:360"
        node5 -->|"Yes"| node6["Clamp color value to 255"]
        click node6 openCode "Modules/Filters_Area.bas:358:360"
        node5 -->|"No"| node7["Keep color value"]
        click node7 openCode "Modules/Filters_Area.bas:353:355"
        node6 --> node8["Write new color to pixel"]
        node7 --> node8
        click node8 openCode "Modules/Filters_Area.bas:363:365"
        node8 --> node9{"Did user cancel?"}
        click node9 openCode "Modules/Filters_Area.bas:369:369"
        node9 -->|"Yes"| node10["Exit filter early"]
        click node10 openCode "Modules/Filters_Area.bas:369:369"
        node9 -->|"No"| node11["Continue to next pixel"]
        click node11 openCode "Modules/Filters_Area.bas:367:372"
        node11 --> node4
    end
    loop2 --> loop3
    loop3 --> node12["Finalize image and finish"]
    click node12 openCode "Modules/Filters_Area.bas:374:380"

classDef HeadingStyle fill:#777777,stroke:#333,stroke-width:2px;
```

<SwmSnippet path="/Modules/Filters_Area.bas" line="277">

---

FilterGridBlur averages vertical and horizontal lines for each pixel, clamps the result, and throttles progress bar updates.

```visual basic
Public Sub FilterGridBlur()

    Message "Generating grids..."

    'Create a local array and point it at the pixel data we want to operate on
    Dim imageData() As Byte, tmpSA As SafeArray2D
    EffectPrep.PrepImageData tmpSA
    workingDIB.WrapArrayAroundDIB imageData, tmpSA
    
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = curDIBValues.Left
    initY = curDIBValues.Top
    finalX = curDIBValues.Right
    finalY = curDIBValues.Bottom
    
    Dim iWidth As Long, iHeight As Long
    iWidth = curDIBValues.Width
    iHeight = curDIBValues.Height
            
    Dim numOfPixels As Long
    numOfPixels = iWidth + iHeight
    
    Dim xStride As Long
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    progBarCheck = ProgressBars.FindBestProgBarValue()
    
    'Finally, a bunch of variables used in color calculation
    Dim r As Long, g As Long, b As Long
    Dim rax() As Long, gax() As Long, bax() As Long
    Dim ray() As Long, gay() As Long, bay() As Long
    ReDim rax(0 To iWidth) As Long, gax(0 To iWidth) As Long, bax(0 To iWidth) As Long
    ReDim ray(0 To iHeight) As Long, gay(0 To iHeight), bay(0 To iHeight)
    
    'Generate the averages for vertical lines
    For x = initX To finalX
        r = 0
        g = 0
        b = 0
        xStride = x * 4
        For y = initY To finalY
            r = r + imageData(xStride + 2, y)
            g = g + imageData(xStride + 1, y)
            b = b + imageData(xStride, y)
        Next y
        rax(x) = r
        gax(x) = g
        bax(x) = b
    Next x
    
    'Generate the averages for horizontal lines
    For y = initY To finalY
        r = 0
        g = 0
        b = 0
        For x = initX To finalX
            xStride = x * 4
            r = r + imageData(xStride + 2, y)
            g = g + imageData(xStride + 1, y)
            b = b + imageData(xStride, y)
        Next x
        ray(y) = r
        gay(y) = g
        bay(y) = b
    Next y
    
    Message "Applying grid blur..."
        
```

---

</SwmSnippet>

<SwmSnippet path="/Modules/Filters_Area.bas" line="347">

---

After prepping the grid blur arrays, we loop through each pixel, average the vertical and horizontal sums, clamp the result, and write it back. Progress bar updates are throttled, and ESC checks let the user cancel if needed.

```visual basic
    'Apply the filter
    For x = initX To finalX
        xStride = x * 4
    For y = initY To finalY
        
        'Average the horizontal and vertical values for each color component
        r = (rax(x) + ray(y)) \ numOfPixels
        g = (gax(x) + gay(y)) \ numOfPixels
        b = (bax(x) + bay(y)) \ numOfPixels
        
        'The colors shouldn't exceed 255, but it doesn't hurt to double-check
        If r > 255 Then r = 255
        If g > 255 Then g = 255
        If b > 255 Then b = 255
        
        'Assign the new RGB values back into the array
        imageData(xStride + 2, y) = r
        imageData(xStride + 1, y) = g
        imageData(xStride, y) = b
        
    Next y
        If (x And progBarCheck) = 0 Then
            If Interface.UserPressedESC() Then Exit For
            SetProgBarVal x
        End If
    Next x
        
```

---

</SwmSnippet>

<SwmSnippet path="/Modules/Filters_Area.bas" line="374">

---

After the pixel loop, FilterGridBlur unwraps the image data array and calls EffectPrep.FinalizeImageData to finish up rendering and cleanup. This makes the changes visible and wraps up the filter.

```visual basic
    'Safely deallocate imageData()
    workingDIB.UnwrapArrayFromDIB imageData
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering
    EffectPrep.FinalizeImageData

End Sub
```

---

</SwmSnippet>

### Distort, Edge, and Plugin Effect Routing

```mermaid
%%{init: {"flowchart": {"defaultRenderer": "elk"}} }%%
flowchart TD
    node1["User selects an effect from the Effects menu"]
    click node1 openCode "Modules/Processor.bas:1723:2014"
    node1 --> node2{"Which effect category is selected?"}
    click node2 openCode "Modules/Processor.bas:1723:2014"
    node2 -->|"Distort, Edge, Light, Noise, Pixelate, Render, Sharpen, Stylize, Transform, Animation, Custom, Plugin"| node3{"Show dialog for user input?"}
    click node3 openCode "Modules/Processor.bas:1723:2014"
    node3 -->|"Yes"| node4["Show effect dialog"]
    click node4 openCode "Modules/Processor.bas:1723:2014"
    node3 -->|"No"| node5["Apply effect directly"]
    click node5 openCode "Modules/Processor.bas:1723:2014"
    node4 --> node6["Effect applied"]
    node5 --> node6
    click node6 openCode "Modules/Processor.bas:1723:2014"

classDef HeadingStyle fill:#777777,stroke:#333,stroke-width:2px;
```

<SwmSnippet path="/Modules/Processor.bas" line="1723">

---

After grid blur, Process_EffectsMenu keeps routing through distort, edge, light, noise, pixelate, render, sharpen, stylize, transform, animation, custom, and plugin effects. For plugins, it hands off to the plugin module, which manages everything from dialogs to execution.

```visual basic
    'Distort filters
    ElseIf Strings.StringsEqual(processID, "Correct lens distortion", True) Then
        If raiseDialog Then ShowPDDialog vbModal, FormLensCorrect Else FormLensCorrect.CorrectLensDistortion processParameters
        Process_EffectsMenu = True
        
    ElseIf Strings.StringsEqual(processID, "Apply lens distortion", True) Then
        If raiseDialog Then ShowPDDialog vbModal, FormLens Else FormLens.ApplyLensDistortion processParameters
        Process_EffectsMenu = True
        
    ElseIf Strings.StringsEqual(processID, "Donut", True) Then
        If raiseDialog Then ShowPDDialog vbModal, FormDonut Else FormDonut.ApplyDonutDistortion processParameters
        Process_EffectsMenu = True
    
    ElseIf Strings.StringsEqual(processID, "Droste", True) Then
        If raiseDialog Then ShowPDDialog vbModal, FormDroste Else FormDroste.FxDroste processParameters
        Process_EffectsMenu = True
        
    ElseIf Strings.StringsEqual(processID, "Miscellaneous distort", True) Then
        If raiseDialog Then ShowPDDialog vbModal, FormMiscDistorts Else FormMiscDistorts.ApplyMiscDistort processParameters
        Process_EffectsMenu = True
        
    ElseIf Strings.StringsEqual(processID, "Pinch and whirl", True) Then
        If raiseDialog Then ShowPDDialog vbModal, FormPinch Else FormPinch.PinchImage processParameters
        Process_EffectsMenu = True
        
    ElseIf Strings.StringsEqual(processID, "Poke", True) Then
        If raiseDialog Then ShowPDDialog vbModal, FormPoke Else FormPoke.ApplyPokeDistort processParameters
        Process_EffectsMenu = True
        
    ElseIf Strings.StringsEqual(processID, "Polar conversion", True) Then
        If raiseDialog Then ShowPDDialog vbModal, FormPolar Else FormPolar.ConvertToPolar processParameters
        Process_EffectsMenu = True
        
    ElseIf Strings.StringsEqual(processID, "Ripple", True) Then
        If raiseDialog Then ShowPDDialog vbModal, FormRipple Else FormRipple.RippleImage processParameters
        Process_EffectsMenu = True
        
    ElseIf Strings.StringsEqual(processID, "Squish", True) Then
        If raiseDialog Then ShowPDDialog vbModal, FormSquish Else FormSquish.SquishImage processParameters
        Process_EffectsMenu = True
        
    ElseIf Strings.StringsEqual(processID, "Swirl", True) Then
        If raiseDialog Then ShowPDDialog vbModal, FormSwirl Else FormSwirl.SwirlImage processParameters
        Process_EffectsMenu = True
        
    ElseIf Strings.StringsEqual(processID, "Waves", True) Then
        If raiseDialog Then ShowPDDialog vbModal, FormWaves Else FormWaves.WaveImage processParameters
        Process_EffectsMenu = True
        
    'Edge filters
    ElseIf Strings.StringsEqual(processID, "Emboss", True) Then
        If raiseDialog Then ShowPDDialog vbModal, FormEmbossEngrave Else FormEmbossEngrave.ApplyEmbossEffect processParameters
        Process_EffectsMenu = True
        
    ElseIf Strings.StringsEqual(processID, "Enhance edges", True) Then
        If raiseDialog Then ShowPDDialog vbModal, FormEdgeEnhance Else FormEdgeEnhance.ApplyEdgeEnhancement processParameters
        Process_EffectsMenu = True
        
    ElseIf Strings.StringsEqual(processID, "Find edges", True) Then
        If raiseDialog Then ShowPDDialog vbModal, FormFindEdges Else FormFindEdges.ApplyEdgeDetection processParameters
        Process_EffectsMenu = True
        
    ElseIf Strings.StringsEqual(processID, "Gradient flow", True) Then
        If raiseDialog Then ShowPDDialog vbModal, FormGradientFlow Else FormGradientFlow.ApplyGradientFlowFx processParameters
        Process_EffectsMenu = True
        
    ElseIf Strings.StringsEqual(processID, "Range filter", True) Then
        If raiseDialog Then ShowPDDialog vbModal, FormRangeFilter Else FormRangeFilter.ApplyRangeFilter processParameters
        Process_EffectsMenu = True
        
    ElseIf Strings.StringsEqual(processID, "Trace contour", True) Then
        If raiseDialog Then ShowPDDialog vbModal, FormContour Else FormContour.TraceContour processParameters
        Process_EffectsMenu = True
    
    'Light and shadow filters
    ElseIf Strings.StringsEqual(processID, "Black light", True) Then
        If raiseDialog Then ShowPDDialog vbModal, FormBlackLight Else FormBlackLight.fxBlackLight processParameters
        Process_EffectsMenu = True
        
    ElseIf Strings.StringsEqual(processID, "Bump map", True) Then
        If raiseDialog Then ShowPDDialog vbModal, FormBumpMap Else FormBumpMap.ApplyBumpMapEffect processParameters
        Process_EffectsMenu = True
        
    ElseIf Strings.StringsEqual(processID, "Cross-screen", True) Then
        If raiseDialog Then ShowPDDialog vbModal, FormCrossScreen Else FormCrossScreen.CrossScreenFilter processParameters
        Process_EffectsMenu = True
        
    ElseIf Strings.StringsEqual(processID, "Rainbow", True) Then
        If raiseDialog Then ShowPDDialog vbModal, FormRainbow Else FormRainbow.ApplyRainbowEffect processParameters
        Process_EffectsMenu = True
        
    ElseIf Strings.StringsEqual(processID, "Sunshine", True) Then
        If raiseDialog Then ShowPDDialog vbModal, FormSunshine Else FormSunshine.fxSunshine processParameters
        Process_EffectsMenu = True
        
    ElseIf Strings.StringsEqual(processID, "Dilate (maximum rank)", True) Then
        If raiseDialog Then Dialogs.PromptEffect_Median 100 Else FormMedian.ApplyMedianFilter processParameters
        Process_EffectsMenu = True
        
    ElseIf Strings.StringsEqual(processID, "Erode (minimum rank)", True) Then
        If raiseDialog Then Dialogs.PromptEffect_Median 1 Else FormMedian.ApplyMedianFilter processParameters
        Process_EffectsMenu = True
        
    'Natural filters
    ElseIf Strings.StringsEqual(processID, "Atmosphere", True) Then
        If raiseDialog Then ShowPDDialog vbModal, FormAtmosphere Else FormAtmosphere.ApplyAtmosphereEffect processParameters
        Process_EffectsMenu = True
        
    ElseIf Strings.StringsEqual(processID, "Fog", True) Then
        If raiseDialog Then ShowPDDialog vbModal, FormFog Else FormFog.fxFog processParameters
        Process_EffectsMenu = True
        
    ElseIf Strings.StringsEqual(processID, "Ignite", True) Then
        If raiseDialog Then ShowPDDialog vbModal, FormIgnite Else FormIgnite.fxBurn processParameters
        Process_EffectsMenu = True
        
    ElseIf Strings.StringsEqual(processID, "Lava", True) Then
        If raiseDialog Then ShowPDDialog vbModal, FormLava Else FormLava.fxLava processParameters
        Process_EffectsMenu = True
        
    ElseIf Strings.StringsEqual(processID, "Metal", True) Then
        If raiseDialog Then ShowPDDialog vbModal, FormMetal Else FormMetal.ApplyMetalFilter processParameters
        Process_EffectsMenu = True
        
    ElseIf Strings.StringsEqual(processID, "Snow", True) Then
        If raiseDialog Then ShowPDDialog vbModal, FormSnow Else FormSnow.ApplySnowEffect processParameters
        Process_EffectsMenu = True
        
    ElseIf Strings.StringsEqual(processID, "Water", True) Then
        If raiseDialog Then ShowPDDialog vbModal, FormWater Else FormWater.ApplyWaterFX processParameters
        Process_EffectsMenu = True
    
    'Noise filters
    ElseIf Strings.StringsEqual(processID, "Add film grain", True) Then
        If raiseDialog Then ShowPDDialog vbModal, FormFilmGrain Else FormFilmGrain.AddFilmGrain processParameters
        Process_EffectsMenu = True
        
    ElseIf Strings.StringsEqual(processID, "Add RGB noise", True) Then
        If raiseDialog Then ShowPDDialog vbModal, FormNoise Else FormNoise.AddNoise processParameters
        Process_EffectsMenu = True
        
    ElseIf Strings.StringsEqual(processID, "Anisotropic diffusion", True) Then
        If raiseDialog Then ShowPDDialog vbModal, FormAnisotropic Else FormAnisotropic.ApplyAnisotropicDiffusion processParameters
        Process_EffectsMenu = True
    
    'Legacy support only; this has been superceded by the new surface blur tool
    ElseIf Strings.StringsEqual(processID, "Bilateral smoothing", True) Then
        If raiseDialog Then ShowPDDialog vbModal, FormSurfaceBlur Else FormSurfaceBlur.BilateralFilter_Central processParameters
        Process_EffectsMenu = True
    
    ElseIf Strings.StringsEqual(processID, "Dust and scratches", True) Then
        If raiseDialog Then ShowPDDialog vbModal, FormDustAndScratches Else FormDustAndScratches.ApplyDustAndScratchesFilter processParameters
        Process_EffectsMenu = True
        
    ElseIf Strings.StringsEqual(processID, "Harmonic mean", True) Then
        If raiseDialog Then ShowPDDialog vbModal, FormHarmonicMean Else FormHarmonicMean.ApplyHarmonicMean processParameters
        Process_EffectsMenu = True
        
    ElseIf Strings.StringsEqual(processID, "Mean shift", True) Then
        If raiseDialog Then ShowPDDialog vbModal, FormMeanShift Else FormMeanShift.ApplyMeanShiftFilter processParameters
        Process_EffectsMenu = True
        
    ElseIf Strings.StringsEqual(processID, "Median", True) Then
        If raiseDialog Then Dialogs.PromptEffect_Median 50 Else FormMedian.ApplyMedianFilter processParameters
        Process_EffectsMenu = True
    
    ElseIf Strings.StringsEqual(processID, "Symmetric nearest-neighbor", True) Then
        If raiseDialog Then ShowPDDialog vbModal, FormSNN Else FormSNN.ApplySymmetricNearestNeighbor processParameters
        Process_EffectsMenu = True
        
    'Pixelate filters
    ElseIf Strings.StringsEqual(processID, "Color halftone", True) Then
        If raiseDialog Then ShowPDDialog vbModal, FormColorHalftone Else FormColorHalftone.ColorHalftoneFilter processParameters
        Process_EffectsMenu = True
        
    ElseIf Strings.StringsEqual(processID, "Crystallize", True) Then
        If raiseDialog Then ShowPDDialog vbModal, FormCrystallize Else FormCrystallize.fxCrystallize processParameters
        Process_EffectsMenu = True
        
    ElseIf Strings.StringsEqual(processID, "Fragment", True) Then
        If raiseDialog Then ShowPDDialog vbModal, FormFragment Else FormFragment.Fragment processParameters
        Process_EffectsMenu = True
        
    ElseIf Strings.StringsEqual(processID, "Mezzotint", True) Then
        If raiseDialog Then ShowPDDialog vbModal, FormMezzotint Else FormMezzotint.ApplyMezzotintEffect processParameters
        Process_EffectsMenu = True
        
    ElseIf Strings.StringsEqual(processID, "Mosaic", True) Then
        If raiseDialog Then ShowPDDialog vbModal, FormMosaic Else FormMosaic.MosaicFilter processParameters
        Process_EffectsMenu = True
    
    ElseIf Strings.StringsEqual(processID, "Pointillize", True) Then
        If raiseDialog Then ShowPDDialog vbModal, FormPointillize Else FormPointillize.Pointillize processParameters
        Process_EffectsMenu = True
        
    'Render filters
    ElseIf Strings.StringsEqual(processID, "Clouds", True) Then
        If raiseDialog Then ShowPDDialog vbModal, FormFxClouds Else FormFxClouds.FxRenderClouds processParameters
        Process_EffectsMenu = True
        
    ElseIf Strings.StringsEqual(processID, "Fibers", True) Then
        If raiseDialog Then ShowPDDialog vbModal, FormFxFibers Else FormFxFibers.FxRenderFibers processParameters
        Process_EffectsMenu = True
    
    ElseIf Strings.StringsEqual(processID, "Truchet", True) Then
        If raiseDialog Then ShowPDDialog vbModal, FormFxTruchet Else FormFxTruchet.FxRenderTruchet processParameters
        Process_EffectsMenu = True
    
    'Sharpen filters
    ElseIf Strings.StringsEqual(processID, "Sharpen", True) Then
        If raiseDialog Then ShowPDDialog vbModal, FormSharpen Else FormSharpen.ApplySharpenFilter processParameters
        Process_EffectsMenu = True
        
    ElseIf Strings.StringsEqual(processID, "Unsharp mask", True) Then
        If raiseDialog Then ShowPDDialog vbModal, FormUnsharpMask Else FormUnsharpMask.UnsharpMask processParameters
        Process_EffectsMenu = True
        
    'Stylize filters
    ElseIf Strings.StringsEqual(processID, "Antique", True) Then
        If raiseDialog Then ShowPDDialog vbModal, FormAntique Else FormAntique.AntiqueEffect processParameters
        Process_EffectsMenu = True
                
    ElseIf Strings.StringsEqual(processID, "Diffuse", True) Then
        If raiseDialog Then ShowPDDialog vbModal, FormDiffuse Else FormDiffuse.DiffuseCustom processParameters
        Process_EffectsMenu = True
        
    ElseIf Strings.StringsEqual(processID, "Kuwahara filter", True) Then
        If raiseDialog Then ShowPDDialog vbModal, FormKuwahara Else FormKuwahara.KuwaharaFilter processParameters
        Process_EffectsMenu = True
        
    ElseIf Strings.StringsEqual(processID, "Outline", True) Then
        If raiseDialog Then ShowPDDialog vbModal, FormOutlineEffect Else FormOutlineEffect.ApplyOutlineEffect processParameters
        Process_EffectsMenu = True
        
    ElseIf Strings.StringsEqual(processID, "Palette", True) Or Strings.StringsEqual(processID, "Palettize", True) Then
        If raiseDialog Then ShowPDDialog vbModal, FormPalettize Else FormPalettize.ApplyPalettizeEffect processParameters
        Process_EffectsMenu = True
        
    ElseIf Strings.StringsEqual(processID, "Portrait glow", True) Then
        If raiseDialog Then ShowPDDialog vbModal, FormPortraitGlow Else FormPortraitGlow.ApplyPortraitGlow processParameters
        Process_EffectsMenu = True
        
    ElseIf Strings.StringsEqual(processID, "Solarize", True) Then
        If raiseDialog Then ShowPDDialog vbModal, FormSolarize Else FormSolarize.SolarizeImage processParameters
        Process_EffectsMenu = True
        
    ElseIf Strings.StringsEqual(processID, "Twins", True) Then
        If raiseDialog Then ShowPDDialog vbModal, FormTwins Else FormTwins.GenerateTwins processParameters
        Process_EffectsMenu = True
        
    ElseIf Strings.StringsEqual(processID, "Vignetting", True) Then
        If raiseDialog Then ShowPDDialog vbModal, FormVignette Else FormVignette.ApplyVignette processParameters
        Process_EffectsMenu = True
        
    'Transform filters
    ElseIf Strings.StringsEqual(processID, "Offset and zoom", True) Then
        If raiseDialog Then ShowPDDialog vbModal, FormPanAndZoom Else FormPanAndZoom.PanAndZoomFilter processParameters
        Process_EffectsMenu = True
        
    ElseIf Strings.StringsEqual(processID, "Perspective", True) Then
        If raiseDialog Then ShowPDDialog vbModal, FormPerspective Else FormPerspective.PerspectiveImage processParameters
        Process_EffectsMenu = True
        
    ElseIf Strings.StringsEqual(processID, "Rotate", True) Then
        If raiseDialog Then ShowPDDialog vbModal, FormRotateDistort Else FormRotateDistort.RotateFilter processParameters
        Process_EffectsMenu = True
        
    ElseIf Strings.StringsEqual(processID, "Shear", True) Then
        If raiseDialog Then ShowPDDialog vbModal, FormShear Else FormShear.ShearImage processParameters
        Process_EffectsMenu = True
        
    ElseIf Strings.StringsEqual(processID, "Spherize", True) Then
        If raiseDialog Then ShowPDDialog vbModal, FormSpherize Else FormSpherize.SpherizeImage processParameters
        Process_EffectsMenu = True
    
    'Animation filters
    ElseIf Strings.StringsEqual(processID, "Animation background", True) Then
        If raiseDialog Then Dialogs.PromptEffect_Animation True Else FormAnimBackground.ApplyAnimationBackground processParameters
        Process_EffectsMenu = True
    
    ElseIf Strings.StringsEqual(processID, "Animation foreground", True) Then
        If raiseDialog Then Dialogs.PromptEffect_Animation False Else FormAnimBackground.ApplyAnimationBackground processParameters
        Process_EffectsMenu = True
    
    ElseIf Strings.StringsEqual(processID, "Animation playback speed", True) Then
        If raiseDialog Then ShowPDDialog vbModal, FormAnimSpeed Else FormAnimSpeed.ApplyNewPlaybackSpeed processParameters
        Process_EffectsMenu = True
        
    'Custom filters
    ElseIf Strings.StringsEqual(processID, "Custom filter", True) Then
        If raiseDialog Then ShowPDDialog vbModal, FormCustomFilter Else Filters_Area.ApplyConvolutionFilter_XML processParameters
        Process_EffectsMenu = True
        
```

---

</SwmSnippet>

<SwmSnippet path="/Modules/Processor.bas" line="2016">

---

At the end of Process_EffectsMenu, we handle plugins by matching the processID and calling the plugin handler if needed. This keeps plugin logic separate and lets the plugin module manage its own workflow.

```visual basic
    '8bf filters have a weird workflow because we simply call "execute" on the plugin but then all handling
    ' occurs inside the plugin - so things like creating Undo data before running an effect doesn't follow a normal workflow.
    ' As such, we use a special module wrapper to handle the details for us.
    ElseIf Strings.StringsEqual(processID, "Photoshop (8bf) plugin", True) Then
        If raiseDialog Then Plugin_8bf.ShowPluginDialog
        Process_EffectsMenu = True
        
    End If
        
End Function
```

---

</SwmSnippet>

## Tools Menu Command Routing

```mermaid
%%{init: {"flowchart": {"defaultRenderer": "elk"}} }%%
flowchart TD
    node1["Receive operation request (processID)"] --> node2{"Is operation handled by tools menu?"}
    click node1 openCode "Modules/Processor.bas:258:258"
    click node2 openCode "Modules/Processor.bas:259:259"
    node2 -->|"Yes"| node3["Perform tool/macro action"]
    click node3 openCode "Modules/Processor.bas:1619:1635"
    node2 -->|"No"| node4{"Is operation a recognized legacy, paint, or macro action?"}
    click node4 openCode "Modules/Processor.bas:262:318"
    node4 -->|"Yes"| node5["Perform legacy/tool/macro action"]
    click node5 openCode "Modules/Processor.bas:262:318"
    node4 -->|"No"| node6["Notify user of unknown operation"]
    click node6 openCode "Modules/Processor.bas:320:327"
    node5 --> node7{"Is this a macro playback for layer/text modification?"}
    click node7 openCode "Modules/Processor.bas:305:318"
    node7 -->|"Yes"| node8["Perform macro layer/text modification"]
    click node8 openCode "Modules/Processor.bas:664:686"
    node7 -->|"No"| node9["Continue with operation"]
    click node9 openCode "Modules/Processor.bas:262:294"
    node3 --> node10{"Did operation modify image?"}
    node8 --> node10
    node9 --> node10
    node10 -->|"Yes"| node11["Update undo/redo and UI"]
    click node11 openCode "Modules/Processor.bas:332:352"
    node10 -->|"No"| node12["No further action"]
    click node12 openCode "Modules/Processor.bas:329:330"
    node11 --> node13{"Should timing be reported? (Undo/Redo created?)"}
    click node13 openCode "Modules/Processor.bas:337:337"
    node13 -->|"Yes"| node14["Report timing"]
    click node14 openCode "Modules/Processor.bas:462:471"
    node13 -->|"No"| node15["Done"]
    node14 --> node15
    node12 --> node15
    node6 --> node15
classDef HeadingStyle fill:#777777,stroke:#333,stroke-width:2px;
```

<SwmSnippet path="/Modules/Processor.bas" line="258">

---

If effects didn't handle the command, Process tries the tools menu next.

```visual basic
    'Tool menu operations have been successfully migrated to XML strings.  (None of their functions raise special return conditions, FYI.)
    If (Not processFound) Then processFound = Process_ToolsMenu(processID, raiseDialog, processParameters, createUndo, relevantTool, recordAction, returnDetails)
    
```

---

</SwmSnippet>

<SwmSnippet path="/Modules/Processor.bas" line="1619">

---

Process_ToolsMenu checks for macro commands—start, stop, play—and calls the relevant Macros method. If the processID doesn't match, it does nothing. This keeps macro logic isolated from other tool commands.

```visual basic
Private Function Process_ToolsMenu(ByVal processID As String, Optional raiseDialog As Boolean = False, Optional processParameters As String = vbNullString, Optional createUndo As PD_UndoType = UNDO_Nothing, Optional relevantTool As Long = -1, Optional recordAction As Boolean = True, Optional ByRef returnDetails As String = vbNullString) As Boolean

    If Strings.StringsEqual(processID, "Start macro recording", True) Then
        Macros.StartMacro
        Process_ToolsMenu = True
        
    ElseIf Strings.StringsEqual(processID, "Stop macro recording", True) Then
        Macros.StopMacro
        Process_ToolsMenu = True
            
    ElseIf Strings.StringsEqual(processID, "Play macro", True) Then
        Macros.PlayMacro
        Process_ToolsMenu = True
        
    End If
        
End Function
```

---

</SwmSnippet>

<SwmSnippet path="/Modules/Processor.bas" line="261">

---

After tools, Process checks for legacy processIDs like paint, fill, and layer/text modifications. For non-destructive changes during macro playback, it calls MiniProcess_NDFX_MacroPlayback to handle those updates cleanly.

```visual basic
    'If the process hasn't been found yet, resume with our legacy processID checks...
    If (Not processFound) Then
    
        'PAINT OPERATIONS
        
        'If we are in the middle of a batch operation, we may actually apply paint strokes in the future.  (This behavior is
        ' currently disabled pending additional testing, however.)  During normal operations, however, we don't need to do
        ' anything here - this processor call just exists to ensure Undo/Redo data was created.
        If Strings.StringsEqual(processID, "Paint stroke", True) Then
            processFound = True
        
        ElseIf Strings.StringsEqual(processID, "Pencil stroke", True) Then
            processFound = True
        
        ElseIf Strings.StringsEqual(processID, "Clone stamp", True) Then
            processFound = True
        
        ElseIf Strings.StringsEqual(processID, "Fill tool", True) Then
            
            'Per https://github.com/tannerhelland/PhotoDemon/issues/286, I'm attempting to support flood fill
            ' operations in recorded macros.  (Apparently this can be a huge timesaver in certain workflows!)
            ' To make this possible, PD needs to know if a macro is currently running; if it is, it will
            ' attempt to manually apply a flood fill.  We do *NOT* want to do this during normal operations,
            ' or it will cause the fill to be applied twice!
            If ((Macros.GetMacroStatus = MacroPLAYBACK) Or (Macros.GetMacroStatus = MacroBATCH)) And (LenB(processParameters) <> 0) Then
                PDImages.GetActiveImage.ResetScratchLayer True
                Tools_Fill.PlayFillFromMacro processParameters
            End If
            
            processFound = True
            
        ElseIf Strings.StringsEqual(processID, "Gradient tool", True) Then
            processFound = True
        
        ElseIf Strings.StringsEqual(processID, "Crop tool", True) Then
            Tools_Crop.Crop_ApplyFromString processParameters
        
        'A "secret" action is used internally by PD when we need some response from the processor engine - like checking for
        ' non-destructive layer changes - but the user is not actually modifying the image.
        ElseIf Strings.StringsEqual(processID, "Do nothing", True) Then
            processFound = True
        
        'Non-destructive layer header modifications are handled by their own specialized non-destructive processor (below).
        ' The only way this case will ever be triggered in *this function* is during macro playback.
        ElseIf Strings.StringsEqualLeft(processID, "Modify layer", True) Then
            If ((Macros.GetMacroStatus = MacroPLAYBACK) Or (Macros.GetMacroStatus = MacroBATCH)) Then
                MiniProcess_NDFX_MacroPlayback thisProcData, False 'Forward the command to a dedicated processor
            End If
            processFound = True
        
        'Text layer modifications are handled by their own specialized non-destructive processor (below).  The only way this case
        ' will ever be triggered is during macro playback.  If encountered, all "modify text layer" instructions follow the same
        ' basic structure: the first parameter is a text setting ID, and the second is a text setting value.
        ElseIf Strings.StringsEqualLeft(processID, "Modify text layer", True) Then
            If ((Macros.GetMacroStatus = MacroPLAYBACK) Or (Macros.GetMacroStatus = MacroBATCH)) Then
                MiniProcess_NDFX_MacroPlayback thisProcData, True 'Forward the command to a dedicated processor
            End If
            processFound = True
                    
```

---

</SwmSnippet>

<SwmSnippet path="/Modules/Processor.bas" line="664">

---

MiniProcess_NDFX_MacroPlayback parses macro parameters to get an action ID and value, then sets the right property on the active layer. If it's a text layer and useTextMode is set, it updates a text property; otherwise, it sets a generic layer property.

```visual basic
Private Sub MiniProcess_NDFX_MacroPlayback(ByRef srcProcData As PD_ProcessCall, Optional ByVal useTextMode As Boolean = False)
    
    'Retrieve any associated parameters from the macro
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    cParams.SetParamString srcProcData.pcParameters
    
    'Action ID may be a generic layer property OR a text layer property (it doesn't matter)
    Dim actionID As Long, defParamValue As Variant
    actionID = cParams.GetLong("id", nameGuaranteedXMLSafe:=True)
    defParamValue = cParams.GetVariant("value")
    
    Dim idxLayer As Long
    idxLayer = PDImages.GetActiveImage.GetActiveLayerIndex
    
    'The layers module handles everything from here
    If (useTextMode And PDImages.GetActiveImage.GetLayerByIndex(idxLayer).IsLayerText()) Then
        PDImages.GetActiveImage.GetActiveLayer.SetTextLayerProperty actionID, defParamValue
    Else
        Layers.SetGenericLayerProperty actionID, defParamValue, idxLayer
    End If
    
End Sub
```

---

</SwmSnippet>

<SwmSnippet path="/Modules/Processor.bas" line="320">

---

After handling all known processIDs, Process checks for unknown ones. If it can't parse the ID, it shows a message box via Interface.bas, prompting the user to report the bug.

```visual basic
        'DEBUG FAILSAFE
        Else
        
            'This function should never be passed a process ID it can't parse, but if that happens,
            ' ask the user to report the unparsed ID
            If (LenB(processID) <> 0) Then PDMsgBox "Unknown processor request submitted: %1" & vbCrLf & vbCrLf & "Please report this bug via the Help -> Submit Bug Report menu.", vbCritical Or vbOKOnly, "Error", processID
            
        End If
        
```

---

</SwmSnippet>

<SwmSnippet path="/Modules/Interface.bas" line="1618">

---

PDMsgBox translates the message and title if needed, replaces any dynamic placeholders, manages the cursor, and tries to show a custom dialog. If that fails, it falls back to the system MsgBox. This keeps error messages readable and localized.

```visual basic
Public Function PDMsgBox(ByVal pMessage As String, ByVal pButtons As VbMsgBoxStyle, ByVal pTitle As String, ParamArray ExtraText() As Variant) As VbMsgBoxResult
    
    'Before passing the message (and any optional parameters) over to the message box dialog, we first need to
    ' plug-in any dynamic elements (e.g. "%n" entries in the message with the param array contents) and apply
    ' any active language translations.
    Dim newMessage As String, newTitle As String
    newMessage = pMessage
    newTitle = pTitle
    
    If (Not g_Language Is Nothing) Then
        If (g_Language.ReadyToTranslate And g_Language.TranslationActive) Then
            newMessage = g_Language.TranslateMessage(pMessage)
            newTitle = g_Language.TranslateMessage(pTitle)
        End If
    End If
    
    'With the message freshly translated, we can plug-in any dynamic text entries
    If (UBound(ExtraText) >= LBound(ExtraText)) Then
        Dim i As Long
        For i = LBound(ExtraText) To UBound(ExtraText)
            newMessage = Replace$(newMessage, "%" & i + 1, CStr(ExtraText(i)))
        Next i
    End If
    
    'Suspend any system-wide cursors, as necessary
    Dim cursorBackup As MousePointerConstants
    cursorBackup = Screen.MousePointer
    Screen.MousePointer = vbDefault
    
    Load dialog_MsgBox
    If dialog_MsgBox.ShowDialog(newMessage, pButtons, newTitle) Then
        PDMsgBox = dialog_MsgBox.DialogResult
    
    'If the dialog failed to load for whatever reason, fall back to a default system message box
    Else
        PDMsgBox = MsgBox(newMessage, pButtons, newTitle)
    End If
    
    'Restore cursor before exiting
    Screen.MousePointer = cursorBackup
    
    Unload dialog_MsgBox
    Set dialog_MsgBox = Nothing

End Function
```

---

</SwmSnippet>

<SwmSnippet path="/Modules/Processor.bas" line="329">

---

After handling all processIDs, Process optionally reports how long the action took if timing is enabled and the action modified the image. This helps with performance tracking and debugging.

```visual basic
    'End of special processID checks
    End If
    
    'Relay any Undo/Redo changes to our processor tracker
    If processFound And (thisProcData.pcUndoType <> createUndo) Then thisProcData.pcUndoType = createUndo
    
    'If the user wants us to time this action, display the results now.  (Note that we only do this for actions that change the image
    ' in some way, as determined by whether meaningful Undo/Redo data is created.)
    If g_DisplayTimingReports And (createUndo <> UNDO_Nothing) Then ReportProcessorTimeTaken m_ProcessingTime
    
    Dim procSortStopTime As Currency
    If (Not raiseDialog) Then VBHacks.GetHighResTime procSortStopTime
    
```

---

</SwmSnippet>

<SwmSnippet path="/Modules/Processor.bas" line="462">

---

ReportProcessorTimeTaken calculates the elapsed time, builds a localized message string, and shows it to the user. This gives immediate feedback on how long the action took.

```visual basic
Public Sub ReportProcessorTimeTaken(ByVal srcStartTime As Currency)

    Dim timingString As String
    timingString = g_Language.TranslateMessage("Time taken")
    timingString = timingString & ": " & Format$(VBHacks.GetTimerDifferenceNow(srcStartTime), "#0.0000") & " "
    timingString = timingString & g_Language.TranslateMessage("seconds")
    
    Message timingString
    
End Sub
```

---

</SwmSnippet>

<SwmSnippet path="/Modules/Processor.bas" line="342">

---

After timing, Process finalizes undo/redo for the action.

```visual basic
    'If the current image has been modified, notify the interface manager of the change.  It will handle things like generating
    ' new thumbnail icons.  (Note that we disable non-essential UI updates while performing batch conversions, to improve performance.)
    If (createUndo <> UNDO_Nothing) And (Macros.GetMacroStatus <> MacroBATCH) Then Interface.NotifyImageChanged PDImages.GetActiveImageID()
    
    Dim procUndoStartTime As Currency
    If (Not raiseDialog) Then VBHacks.GetHighResTime procUndoStartTime
    
    'After an action completes, figure out if we need to push a new entry onto the Undo/Redo stack.  (Note that for convenience,
    ' this sub also handles roll-back of some UI elements if the current operation was canceled prematurely.)
    FinalizeUndoRedoState thisProcData, PDImages.GetActiveImage
    
    Dim procUndoStopTime As Currency
    If (Not raiseDialog) Then VBHacks.GetHighResTime procUndoStopTime
    
```

---

</SwmSnippet>

## Undo/Redo Finalization and Cancellation Handling

```mermaid
%%{init: {"flowchart": {"defaultRenderer": "elk"}} }%%
flowchart TD
    node1["Finalize undo/redo state after user action"] --> node2{"Was the action canceled? (cancel trigger)"}
    click node1 openCode "Modules/Processor.bas:1290:1292"
    click node2 openCode "Modules/Processor.bas:1293:1294"
    node2 -->|"Yes"| node3["Reset interface and cancel trigger"]
    click node3 openCode "Modules/Processor.bas:1295:1300"
    node2 -->|"No"| node4{"Eligible for undo/redo? (undo type, macro status, recorded status)"}
    click node4 openCode "Modules/Processor.bas:1306:1312"
    node4 -->|"Yes"| node5{"Is action 'Fade'?"}
    click node5 openCode "Modules/Processor.bas:1329:1332"
    node5 -->|"Yes"| node7["Handle 'Fade' layer for undo"]
    click node7 openCode "Modules/Processor.bas:1330:1332"
    node5 -->|"No"| node6["Record current image state for undo/redo"]
    click node6 openCode "Modules/Processor.bas:1334:1335"
    node7 --> node6
    node4 -->|"No"| node8["No undo/redo entry created"]
    click node8 openCode "Modules/Processor.bas:1338:1340"
classDef HeadingStyle fill:#777777,stroke:#333,stroke-width:2px;
```

<SwmSnippet path="/Modules/Processor.bas" line="1290">

---

In FinalizeUndoRedoState, if the user cancels the action, we reset progress bars, show a cancellation message, and clear the cancel flag. This keeps the UI and workflow clean after a canceled operation.

```visual basic
Private Sub FinalizeUndoRedoState(ByRef srcProcData As PD_ProcessCall, ByRef targetImage As pdImage)

    'If the user canceled the requested action before it completed, we may need to manually roll back some processor phases
    If g_cancelCurrentAction Then
        
        'Reset any interface elements that may still be in "processing" mode
        ProgressBars.ReleaseProgressBar
        Message "Action canceled."
    
```

---

</SwmSnippet>

<SwmSnippet path="/Modules/Processor.bas" line="1299">

---

After handling cancellation, FinalizeUndoRedoState checks if undo data should be created. It picks the right layer (handling selection and fade cases), then calls the UndoManager to create the entry. Undo/redo is skipped during macro playback.

```visual basic
        'Reset the cancel trigger; if this is not done, the user will not be able to cancel subsequent actions.
        g_cancelCurrentAction = False
        
    'If the user did not cancel the current process request, and this request modified the image, push the current image state
    ' onto the Undo/Redo stack
    Else
    
        'Generally, we assume that actions want us to create Undo data for them.  However, there are a few known exceptions:
        ' 1) If this processor request was a UI-only action (e.g. displaying a dialog)
        ' 2) If macro recording has been forcibly disabled.  (This is typically used when an internal PD function
        '     utilizes other functions, but we only want a single Undo point created for the full set of actions.)
        ' 3) If we are in the midst of playing back a recorded macro.  (Undo/Redo entries take time and memory to process,
        '     so we ignore them during macro playback)
        If (srcProcData.pcUndoType <> UNDO_Nothing) And (Macros.GetMacroStatus <> MacroBATCH) And srcProcData.pcRecorded Then
            If PDImages.IsImageActive() Then
                
                'In most cases, the Undo/Redo engine can automatically figure out what layer is affected by the current action.
                ' In some rare cases, however, the affected portion of the image may not be obvious.
                '
                'Let's start by grabbing the active layer ID.  (We cache it in case subsequent modifications cause it to change.)
                Dim affectedLayerID As Long
                If (srcProcData.pcUndoType = UNDO_Selection) Then
                    affectedLayerID = -1
                Else
                    affectedLayerID = targetImage.GetActiveLayerID
                End If
                
                'The "Edit > Fade" action is unique, because it does not necessarily affect the active layer (e.g. if the user blurs
                ' a layer, then switches to a new layer, Fade will affect the *old layer* only).  Find the relevant layer ID
                ' before calling the Undo engine.
                If Strings.StringsEqual(srcProcData.pcID, "Fade", True) Then
                    Dim tmpDIB As pdDIB
                    targetImage.UndoManager.FillDIBWithLastUndoCopy tmpDIB, affectedLayerID, , True
                End If
            
                'Create the Undo data
                targetImage.UndoManager.CreateUndoData srcProcData, affectedLayerID
            
            End If
        End If
    
    End If
    
End Sub
```

---

</SwmSnippet>

## UI State Restoration and Final Sync

```mermaid
%%{init: {"flowchart": {"defaultRenderer": "elk"}} }%%
flowchart TD
    node1["Finalize UI updates after processing"] --> node2{"Was undo action performed?"}
    click node1 openCode "Modules/Processor.bas:356:359"
    node2 -->|"Yes"| node3["Return focus to active image"]
    click node2 openCode "Modules/Processor.bas:362:363"
    node2 -->|"No"| node4{"Is macro batch processing active?"}
    click node4 openCode "Modules/Processor.bas:370:383"
    node4 -->|"Yes"| node8["Restore main form and controls"]
    node4 -->|"No"| node5{"Was dialog raised?"}
    click node5 openCode "Modules/Processor.bas:373:376"
    node5 -->|"Yes"| node6{"Was dialog canceled?"}
    click node6 openCode "Modules/Processor.bas:376:377"
    node6 -->|"No"| node7["Sync UI to current image"]
    click node7 openCode "Modules/Processor.bas:376:377"
    node6 -->|"Yes"| node8["Restore main form and controls"]
    node5 -->|"No"| node7
    node3 --> node8["Restore main form and controls"]
    click node8 openCode "Modules/Processor.bas:389:389"
    node7 --> node8
    node8 --> node9{"Is update available?"}
    click node9 openCode "Modules/Processor.bas:392:392"
    node9 -->|"Yes"| node10["Display update notification"]
    click node10 openCode "Modules/Processor.bas:392:392"
    node9 -->|"No"| node11["Log process timing"]
    node10 --> node11
    click node11 openCode "Modules/Processor.bas:395:396"
    node11 --> node12{"Did error occur?"}
    click node12 openCode "Modules/Processor.bas:402:440"
    node12 -->|"No"| node19["End"]
    node12 -->|"Yes"| node13{"Error type?"}
    node13 -->|"Error number 0"| node15["Ignore error"]
    click node15 openCode "Modules/Processor.bas:416:419"
    node13 -->|"Object unloaded"| node16["Ignore error"]
    click node16 openCode "Modules/Processor.bas:422:425"
    node13 -->|"Out of memory"| node14["Show out of memory message"]
    click node14 openCode "Modules/Processor.bas:428:431"
    node13 -->|"Unknown error"| node17["Show unknown error message"]
    click node17 openCode "Modules/Processor.bas:435:439"
    node17 --> node18{"User agrees to file bug report?"}
    click node18 openCode "Modules/Processor.bas:446:447"
    node18 -->|"Yes"| node20["Guide user to bug report"]
    click node20 openCode "Modules/Processor.bas:1345:1353"
    node18 -->|"No"| node19["End"]
    click node19 openCode "Modules/Processor.bas:398:398"
classDef HeadingStyle fill:#777777,stroke:#333,stroke-width:2px;
```

<SwmSnippet path="/Modules/Processor.bas" line="356">

---

After undo/redo, Process restores the UI to idle, re-enables user interaction, and syncs the interface if needed. This keeps the app responsive and up-to-date after any action.

```visual basic
    'From this point onward, we're only going to be finalizing UI updates.  Some of these updates will not trigger
    ' if the central processor is active (by design, to avoid excessive redraws), so to ensure that they trigger *now*,
    ' we need to mark the processor as "idle".
    m_Processing = False
    
    'If a filter or tool was just used, return focus to the active form.  This will make it "flash" to catch the user's attention.
    If (createUndo <> UNDO_Nothing) Then
        If PDImages.IsImageActive() Then CanvasManager.ActivatePDImage PDImages.GetActiveImageID(), "processor call complete", True, createUndo
    
    'The interface will automatically be synched if an image is open and some undo-related action was applied (via the
    ' ActivatePDImage function, above).  If an undo-related action was *not* applied, it's harder to know if an interface
    ' sync is required.  Run some tests to see if we can skip this step.
    Else
        
        If (Macros.GetMacroStatus <> MacroBATCH) Then
            
            'If a dialog was raised via PD's raiseDialog function, we may be able to skip a UI sync
            If raiseDialog Then
            
                'If the raised dialog was canceled, skip a UI sync entirely, as nothing has changed
                If Not (Interface.GetLastShowDialogResult = vbCancel) Then Interface.SyncInterfaceToCurrentImage
            
            'If no dialog was shown, a resync is required as we can't guarantee that the image state is unchanged
            Else
                Interface.SyncInterfaceToCurrentImage
            End If
            
        End If
        
    End If
    
    'Re-enable the main form and restore things like selection animations and proper control focus.
    ' (NOTE: this call is also what decrements the nested process counter.)
    SetProcessorUI_Idle processID, raiseDialog, processParameters, createUndo, relevantTool, recordAction
    
    'PD periodically checks for background updates.  If one is available, and we haven't displayed a notification yet, do so now
    If Updates.IsUpdateReadyToInstall() Then Updates.DisplayUpdateNotification
    
    Dim procFinalStopTime As Currency
    If (Not raiseDialog) Then VBHacks.GetHighResTime procFinalStopTime
    If (Not raiseDialog) Then PDDebug.LogAction "Net time for """ & processID & """: " & VBHacks.GetTimeDiffAsString(procStartTime, procFinalStopTime) & ".  (init: " & VBHacks.GetTimeDiffAsString(procStartTime, procSortStartTime) & ", sort: " & VBHacks.GetTimeDiffAsString(procSortStartTime, procSortStopTime) & ", pre-Undo: " & VBHacks.GetTimeDiffAsString(procSortStopTime, procUndoStartTime) & ", undo: " & VBHacks.GetTimeDiffAsString(procUndoStartTime, procUndoStopTime) & ", UI: " & VBHacks.GetTimeDiffAsString(procUndoStopTime, procFinalStopTime) & ")"
    
    Exit Sub

'MAIN PHOTODEMON ERROR HANDLER STARTS HERE

MainErrHandler:
    
    PDDebug.LogAction "WARNING: Processor module had an error (" & Err.Number & "): " & Err.Description
    
    'Re-enable the main form and restore things like selection animations and proper control focus
    SetProcessorUI_Idle processID, raiseDialog, processParameters, createUndo, relevantTool, recordAction
    
```

---

</SwmSnippet>

<SwmSnippet path="/Modules/Processor.bas" line="409">

---

After restoring the UI, Process flushes any pending UI syncs and handles errors by showing localized messages. This keeps the interface consistent and gives users clear feedback if something goes wrong.

```visual basic
    'Ensure any pending UI syncs are flushed
    Interface.SyncInterfaceToCurrentImage

    'Attempt to generate a human-readable error message
    Dim addInfo As String, mType As VbMsgBoxStyle, msgReturn As VbMsgBoxResult
    
    'Ignore errors that aren't actually errors
    If (Err.Number = 0) Then
        Err.Clear
        On Error GoTo 0
        Exit Sub
    
    'Object was unloaded before it could be shown - this is intentional, so ignore the error
    ElseIf (Err.Number = 364) Then
        Err.Clear
        On Error GoTo 0
        Exit Sub
    
    'Out of memory error
    ElseIf ((Err.Number = 480) Or (Err.Number = 7)) Then
        On Error GoTo 0
        addInfo = g_Language.TranslateMessage("There is not enough memory available to continue this operation.  Please free up system memory (RAM) by shutting down unneeded programs - especially your web browser, if it is open - then try the action again.")
        Message "Out of memory.  Function canceled."
        mType = vbExclamation Or vbOKOnly
        
    'Unknown error
    Else
        On Error GoTo 0
        addInfo = g_Language.TranslateMessage("PhotoDemon cannot locate additional information for this error.  That probably means this error is a bug, and it needs to be fixed!" & vbCrLf & vbCrLf & "Would you like to submit a bug report?  (It takes less than one minute, and it helps everyone who uses the software.)")
        mType = vbCritical Or vbYesNo
        Message "Unknown error."
    End If
    
```

---

</SwmSnippet>

<SwmSnippet path="/Modules/Processor.bas" line="442">

---

After coming back from Interface.bas, Process uses PDMsgBox to show a detailed, localized error message with all the relevant info.

```visual basic
    'Create the message box to return the error information
    msgReturn = PDMsgBox("PhotoDemon has experienced an error.  Details on the problem include:" & vbCrLf & vbCrLf & "Error number %1" & vbCrLf & "Description: %2" & vbCrLf & vbCrLf & "%3", mType, "Error", Err.Number, Err.Description, addInfo)
    
```

---

</SwmSnippet>

<SwmSnippet path="/Modules/Processor.bas" line="445">

---

After returning from Interface.bas, Process checks the user's response to the error dialog. If they agree to file a bug report, it calls FileErrorReport with the error number, handing off the reporting flow.

```visual basic
    'If the message box return value is "Yes", the user is willing to file a bug report.
    If (msgReturn = vbYes) Then FileErrorReport Err.Number
        
End Sub
```

---

</SwmSnippet>

<SwmSnippet path="/Modules/Processor.bas" line="1345">

---

FileErrorReport opens the GitHub issues page in the user's browser so they can report a bug directly. Right after, it calls PDMsgBox from Interface.bas to show instructions, telling the user to include the error number in their report for easier debugging.

```visual basic
Private Sub FileErrorReport(ByVal errNumber As Long)

    'Shell a browser window with the GitHub issue report form
    Web.OpenURL "https://github.com/tannerhelland/PhotoDemon/issues/"
    
    'Display one final message box with additional instructions
    PDMsgBox "PhotoDemon has automatically opened the bug report webpage for you.  Please click the ""New Issue"" button, then select ""Bug Report"".  Answer the questions as best you can, and please include the following error number somewhere in your report: %1" & vbCrLf & vbCrLf & "When finished, click the Submit New Issue button.  Thank you!", vbInformation Or vbOKOnly, "Bug report instructions", errNumber
    
End Sub
```

---

</SwmSnippet>

&nbsp;

*This is an auto-generated document by Swimm 🌊 and has not yet been verified by a human*

<SwmMeta version="3.0.0" repo-id="Z2l0aHViJTNBJTNBVkI2LVBob3RvRGVtb24lM0ElM0FTd2ltbS1EZW1v" repo-name="VB6-PhotoDemon"><sup>Powered by [Swimm](https://app.swimm.io/)</sup></SwmMeta>
