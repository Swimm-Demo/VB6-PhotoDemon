---
title: Processing image operations
---
This document describes how user actions—such as selecting a menu item, running a macro, or applying an effect—are processed and executed. The flow covers both interactive and automated actions, supporting <SwmToken path="Modules/Interface.bas" pos="278:33:35" line-data="            &#39; (like &quot;Flatten&quot; or &quot;Delete layer&quot;) are only relevant if this is a multi-layer image.">`multi-layer`</SwmToken> and advanced editing features. Each operation is dispatched to the appropriate handler, executed, and finalized with UI updates and undo/redo management.

```mermaid
flowchart TD
  node1["Processing Entry and Pre-Operation Checks"]:::HeadingStyle
  click node1 goToHeading "Processing Entry and Pre-Operation Checks"
  node1 --> node2["File Menu Command Dispatch and Dialog Handling"]:::HeadingStyle
  click node2 goToHeading "File Menu Command Dispatch and Dialog Handling"
  node2 --> node3["Edit Menu Command Handling"]:::HeadingStyle
  click node3 goToHeading "Edit Menu Command Handling"
  node3 --> node4["Image Menu Command Handling and Dialog Routing"]:::HeadingStyle
  click node4 goToHeading "Image Menu Command Handling and Dialog Routing"
  node4 --> node5["Layer Operations Dispatch"]:::HeadingStyle
  click node5 goToHeading "Layer Operations Dispatch"
  node5 --> node6["Selection Operations Dispatch"]:::HeadingStyle
  click node6 goToHeading "Selection Operations Dispatch"
  node6 --> node7["Image Adjustments Dispatch"]:::HeadingStyle
  click node7 goToHeading "Image Adjustments Dispatch"
  node7 --> node8["Effect Processing and Dialog Routing"]:::HeadingStyle
  click node8 goToHeading "Effect Processing and Dialog Routing"
  node8 --> node9["Tool Menu Command Dispatch and Macro Operations"]:::HeadingStyle
  click node9 goToHeading "Tool Menu Command Dispatch and Macro Operations"
classDef HeadingStyle fill:#777777,stroke:#333,stroke-width:2px;
```

# Where is this flow used?

This flow is used multiple times in the codebase as represented in the following diagram:

(Note - these are only some of the entry points of this flow)

```mermaid
graph TD;
      cf65582bc3e0e529fa2c2d8383f195bac1835c2925c0647c45307471ab620ba1(Modules/Actions.bas::LaunchAction_ByName) --> 84f8cb768dd9c6f851e0871b9090b2ff372db320a69fdbb4aadde9f2fa8c107c(Modules/Actions.bas::Launch_ByName_MenuTools)

cf65582bc3e0e529fa2c2d8383f195bac1835c2925c0647c45307471ab620ba1(Modules/Actions.bas::LaunchAction_ByName) --> 12c9a6e776b714d65778b80231418fd238e516829778991c92dc834f649dc498(Modules/Actions.bas::Launch_ByName_MenuFile)

cf65582bc3e0e529fa2c2d8383f195bac1835c2925c0647c45307471ab620ba1(Modules/Actions.bas::LaunchAction_ByName) --> 738ae8b17c0e968e2f36bb0b907feda2d4589c2d81ef00a6f7d5196daa6a8203(Modules/Actions.bas::Launch_ByName_MenuEdit)

cf65582bc3e0e529fa2c2d8383f195bac1835c2925c0647c45307471ab620ba1(Modules/Actions.bas::LaunchAction_ByName) --> 98f62c390e547e9c34e8d51bf26b6531faa26a29caf30d148f1b51cd27d0a2cf(Modules/Actions.bas::Launch_ByName_MenuImage)

cf65582bc3e0e529fa2c2d8383f195bac1835c2925c0647c45307471ab620ba1(Modules/Actions.bas::LaunchAction_ByName) --> 45c8fe9f692f2ef627e4b51417c7a8726c1fc30478e11a00decacecb68cb170d(Modules/Actions.bas::Launch_ByName_MenuLayer)

cf65582bc3e0e529fa2c2d8383f195bac1835c2925c0647c45307471ab620ba1(Modules/Actions.bas::LaunchAction_ByName) --> cdd4c58891a4902decc4563edbf0d265f688cb7491c4b5ee3eeb3d150483d5aa(Modules/Actions.bas::Launch_ByName_MenuSelect)

cf65582bc3e0e529fa2c2d8383f195bac1835c2925c0647c45307471ab620ba1(Modules/Actions.bas::LaunchAction_ByName) --> 4c1cfc4c9608f89908cf19a9062aaaf78d00984d2f580c68be0df07fa8c8f0de(Modules/Actions.bas::Launch_ByName_MenuAdjustments)

cf65582bc3e0e529fa2c2d8383f195bac1835c2925c0647c45307471ab620ba1(Modules/Actions.bas::LaunchAction_ByName) --> 0de266da5af536af26b16eb40445e1e8ae1e9c63f92d9bb7967e72b1f63e4e5e(Modules/Actions.bas::Launch_ByName_MenuEffects)

84f8cb768dd9c6f851e0871b9090b2ff372db320a69fdbb4aadde9f2fa8c107c(Modules/Actions.bas::Launch_ByName_MenuTools) --> 519e290ec0b7ab0348806e0f4d64e2af2d9fc7f8f0baef99518dd681638ee3a0(Modules/Processor.bas::Process):::mainFlowStyle

12c9a6e776b714d65778b80231418fd238e516829778991c92dc834f649dc498(Modules/Actions.bas::Launch_ByName_MenuFile) --> 519e290ec0b7ab0348806e0f4d64e2af2d9fc7f8f0baef99518dd681638ee3a0(Modules/Processor.bas::Process):::mainFlowStyle

738ae8b17c0e968e2f36bb0b907feda2d4589c2d81ef00a6f7d5196daa6a8203(Modules/Actions.bas::Launch_ByName_MenuEdit) --> 519e290ec0b7ab0348806e0f4d64e2af2d9fc7f8f0baef99518dd681638ee3a0(Modules/Processor.bas::Process):::mainFlowStyle

98f62c390e547e9c34e8d51bf26b6531faa26a29caf30d148f1b51cd27d0a2cf(Modules/Actions.bas::Launch_ByName_MenuImage) --> 519e290ec0b7ab0348806e0f4d64e2af2d9fc7f8f0baef99518dd681638ee3a0(Modules/Processor.bas::Process):::mainFlowStyle

45c8fe9f692f2ef627e4b51417c7a8726c1fc30478e11a00decacecb68cb170d(Modules/Actions.bas::Launch_ByName_MenuLayer) --> 519e290ec0b7ab0348806e0f4d64e2af2d9fc7f8f0baef99518dd681638ee3a0(Modules/Processor.bas::Process):::mainFlowStyle

cdd4c58891a4902decc4563edbf0d265f688cb7491c4b5ee3eeb3d150483d5aa(Modules/Actions.bas::Launch_ByName_MenuSelect) --> 519e290ec0b7ab0348806e0f4d64e2af2d9fc7f8f0baef99518dd681638ee3a0(Modules/Processor.bas::Process):::mainFlowStyle

4c1cfc4c9608f89908cf19a9062aaaf78d00984d2f580c68be0df07fa8c8f0de(Modules/Actions.bas::Launch_ByName_MenuAdjustments) --> 519e290ec0b7ab0348806e0f4d64e2af2d9fc7f8f0baef99518dd681638ee3a0(Modules/Processor.bas::Process):::mainFlowStyle

0de266da5af536af26b16eb40445e1e8ae1e9c63f92d9bb7967e72b1f63e4e5e(Modules/Actions.bas::Launch_ByName_MenuEffects) --> 519e290ec0b7ab0348806e0f4d64e2af2d9fc7f8f0baef99518dd681638ee3a0(Modules/Processor.bas::Process):::mainFlowStyle

2f11e6b3d753e6c408055d6003e8dbb9a0efbee9bb854de86a4e972a500dd2bd(Modules/Layers.bas::LoadImageAsNewLayer) --> 519e290ec0b7ab0348806e0f4d64e2af2d9fc7f8f0baef99518dd681638ee3a0(Modules/Processor.bas::Process):::mainFlowStyle

e0eb7a05ee3d732a3e030589917f9ea100a3a88134e1b8f5c0187f7264b31f1f(Modules/MoveTool.bas::NotifyKeyDown) --> 519e290ec0b7ab0348806e0f4d64e2af2d9fc7f8f0baef99518dd681638ee3a0(Modules/Processor.bas::Process):::mainFlowStyle

19b3e55ebbcb2a6666ca9cae9018fcf945a202eb141c649ad360112bde8cce33(Modules/SelectionUI.bas::NotifySelectionMouseUp) --> 519e290ec0b7ab0348806e0f4d64e2af2d9fc7f8f0baef99518dd681638ee3a0(Modules/Processor.bas::Process):::mainFlowStyle

3c23c9787d44bbfc0a8bcfa6c1ef2949a29b511834fc1ef2cdb87250aebbc28e(Forms/File_BatchWizard.frm::cmdNext_Click) --> 9f5e49590f5e5ed922a86696731778583de59723fc9d1ad4ade858bfbe6f9dae(Forms/File_BatchWizard.frm::ChangeBatchPage)

9f5e49590f5e5ed922a86696731778583de59723fc9d1ad4ade858bfbe6f9dae(Forms/File_BatchWizard.frm::ChangeBatchPage) --> 0bc4bcdf382f08f7d35863edc83db6d90281bde20daf6b2544d965b2a3ae83f4(Forms/File_BatchWizard.frm::PrepareForBatchConversion)

0bc4bcdf382f08f7d35863edc83db6d90281bde20daf6b2544d965b2a3ae83f4(Forms/File_BatchWizard.frm::PrepareForBatchConversion) --> 822dbdc46402979c30729256f4ee31d1b3ae13f2f96487565238e6b45c2c39f2(Forms/File_BatchWizard.frm::ApplyEditOperations)

822dbdc46402979c30729256f4ee31d1b3ae13f2f96487565238e6b45c2c39f2(Forms/File_BatchWizard.frm::ApplyEditOperations) --> 519e290ec0b7ab0348806e0f4d64e2af2d9fc7f8f0baef99518dd681638ee3a0(Modules/Processor.bas::Process):::mainFlowStyle


classDef mainFlowStyle color:#000000,fill:#7CB9F4
classDef rootsStyle color:#000000,fill:#00FFF4
classDef Style1 color:#000000,fill:#00FFAA
classDef Style2 color:#000000,fill:#FFFF00
classDef Style3 color:#000000,fill:#AA7CB9

%% Swimm:
%% graph TD;
%%       cf65582bc3e0e529fa2c2d8383f195bac1835c2925c0647c45307471ab620ba1(<SwmPath>[Modules/Actions.bas](Modules/Actions.bas)</SwmPath>::LaunchAction_ByName) --> 84f8cb768dd9c6f851e0871b9090b2ff372db320a69fdbb4aadde9f2fa8c107c(<SwmPath>[Modules/Actions.bas](Modules/Actions.bas)</SwmPath>::Launch_ByName_MenuTools)
%% 
%% cf65582bc3e0e529fa2c2d8383f195bac1835c2925c0647c45307471ab620ba1(<SwmPath>[Modules/Actions.bas](Modules/Actions.bas)</SwmPath>::LaunchAction_ByName) --> 12c9a6e776b714d65778b80231418fd238e516829778991c92dc834f649dc498(<SwmPath>[Modules/Actions.bas](Modules/Actions.bas)</SwmPath>::Launch_ByName_MenuFile)
%% 
%% cf65582bc3e0e529fa2c2d8383f195bac1835c2925c0647c45307471ab620ba1(<SwmPath>[Modules/Actions.bas](Modules/Actions.bas)</SwmPath>::LaunchAction_ByName) --> 738ae8b17c0e968e2f36bb0b907feda2d4589c2d81ef00a6f7d5196daa6a8203(<SwmPath>[Modules/Actions.bas](Modules/Actions.bas)</SwmPath>::Launch_ByName_MenuEdit)
%% 
%% cf65582bc3e0e529fa2c2d8383f195bac1835c2925c0647c45307471ab620ba1(<SwmPath>[Modules/Actions.bas](Modules/Actions.bas)</SwmPath>::LaunchAction_ByName) --> 98f62c390e547e9c34e8d51bf26b6531faa26a29caf30d148f1b51cd27d0a2cf(<SwmPath>[Modules/Actions.bas](Modules/Actions.bas)</SwmPath>::Launch_ByName_MenuImage)
%% 
%% cf65582bc3e0e529fa2c2d8383f195bac1835c2925c0647c45307471ab620ba1(<SwmPath>[Modules/Actions.bas](Modules/Actions.bas)</SwmPath>::LaunchAction_ByName) --> 45c8fe9f692f2ef627e4b51417c7a8726c1fc30478e11a00decacecb68cb170d(<SwmPath>[Modules/Actions.bas](Modules/Actions.bas)</SwmPath>::Launch_ByName_MenuLayer)
%% 
%% cf65582bc3e0e529fa2c2d8383f195bac1835c2925c0647c45307471ab620ba1(<SwmPath>[Modules/Actions.bas](Modules/Actions.bas)</SwmPath>::LaunchAction_ByName) --> cdd4c58891a4902decc4563edbf0d265f688cb7491c4b5ee3eeb3d150483d5aa(<SwmPath>[Modules/Actions.bas](Modules/Actions.bas)</SwmPath>::Launch_ByName_MenuSelect)
%% 
%% cf65582bc3e0e529fa2c2d8383f195bac1835c2925c0647c45307471ab620ba1(<SwmPath>[Modules/Actions.bas](Modules/Actions.bas)</SwmPath>::LaunchAction_ByName) --> 4c1cfc4c9608f89908cf19a9062aaaf78d00984d2f580c68be0df07fa8c8f0de(<SwmPath>[Modules/Actions.bas](Modules/Actions.bas)</SwmPath>::Launch_ByName_MenuAdjustments)
%% 
%% cf65582bc3e0e529fa2c2d8383f195bac1835c2925c0647c45307471ab620ba1(<SwmPath>[Modules/Actions.bas](Modules/Actions.bas)</SwmPath>::LaunchAction_ByName) --> 0de266da5af536af26b16eb40445e1e8ae1e9c63f92d9bb7967e72b1f63e4e5e(<SwmPath>[Modules/Actions.bas](Modules/Actions.bas)</SwmPath>::Launch_ByName_MenuEffects)
%% 
%% 84f8cb768dd9c6f851e0871b9090b2ff372db320a69fdbb4aadde9f2fa8c107c(<SwmPath>[Modules/Actions.bas](Modules/Actions.bas)</SwmPath>::Launch_ByName_MenuTools) --> 519e290ec0b7ab0348806e0f4d64e2af2d9fc7f8f0baef99518dd681638ee3a0(<SwmPath>[Modules/Processor.bas](Modules/Processor.bas)</SwmPath>::Process):::mainFlowStyle
%% 
%% 12c9a6e776b714d65778b80231418fd238e516829778991c92dc834f649dc498(<SwmPath>[Modules/Actions.bas](Modules/Actions.bas)</SwmPath>::Launch_ByName_MenuFile) --> 519e290ec0b7ab0348806e0f4d64e2af2d9fc7f8f0baef99518dd681638ee3a0(<SwmPath>[Modules/Processor.bas](Modules/Processor.bas)</SwmPath>::Process):::mainFlowStyle
%% 
%% 738ae8b17c0e968e2f36bb0b907feda2d4589c2d81ef00a6f7d5196daa6a8203(<SwmPath>[Modules/Actions.bas](Modules/Actions.bas)</SwmPath>::Launch_ByName_MenuEdit) --> 519e290ec0b7ab0348806e0f4d64e2af2d9fc7f8f0baef99518dd681638ee3a0(<SwmPath>[Modules/Processor.bas](Modules/Processor.bas)</SwmPath>::Process):::mainFlowStyle
%% 
%% 98f62c390e547e9c34e8d51bf26b6531faa26a29caf30d148f1b51cd27d0a2cf(<SwmPath>[Modules/Actions.bas](Modules/Actions.bas)</SwmPath>::Launch_ByName_MenuImage) --> 519e290ec0b7ab0348806e0f4d64e2af2d9fc7f8f0baef99518dd681638ee3a0(<SwmPath>[Modules/Processor.bas](Modules/Processor.bas)</SwmPath>::Process):::mainFlowStyle
%% 
%% 45c8fe9f692f2ef627e4b51417c7a8726c1fc30478e11a00decacecb68cb170d(<SwmPath>[Modules/Actions.bas](Modules/Actions.bas)</SwmPath>::Launch_ByName_MenuLayer) --> 519e290ec0b7ab0348806e0f4d64e2af2d9fc7f8f0baef99518dd681638ee3a0(<SwmPath>[Modules/Processor.bas](Modules/Processor.bas)</SwmPath>::Process):::mainFlowStyle
%% 
%% cdd4c58891a4902decc4563edbf0d265f688cb7491c4b5ee3eeb3d150483d5aa(<SwmPath>[Modules/Actions.bas](Modules/Actions.bas)</SwmPath>::Launch_ByName_MenuSelect) --> 519e290ec0b7ab0348806e0f4d64e2af2d9fc7f8f0baef99518dd681638ee3a0(<SwmPath>[Modules/Processor.bas](Modules/Processor.bas)</SwmPath>::Process):::mainFlowStyle
%% 
%% 4c1cfc4c9608f89908cf19a9062aaaf78d00984d2f580c68be0df07fa8c8f0de(<SwmPath>[Modules/Actions.bas](Modules/Actions.bas)</SwmPath>::Launch_ByName_MenuAdjustments) --> 519e290ec0b7ab0348806e0f4d64e2af2d9fc7f8f0baef99518dd681638ee3a0(<SwmPath>[Modules/Processor.bas](Modules/Processor.bas)</SwmPath>::Process):::mainFlowStyle
%% 
%% 0de266da5af536af26b16eb40445e1e8ae1e9c63f92d9bb7967e72b1f63e4e5e(<SwmPath>[Modules/Actions.bas](Modules/Actions.bas)</SwmPath>::Launch_ByName_MenuEffects) --> 519e290ec0b7ab0348806e0f4d64e2af2d9fc7f8f0baef99518dd681638ee3a0(<SwmPath>[Modules/Processor.bas](Modules/Processor.bas)</SwmPath>::Process):::mainFlowStyle
%% 
%% 2f11e6b3d753e6c408055d6003e8dbb9a0efbee9bb854de86a4e972a500dd2bd(<SwmPath>[Modules/Layers.bas](Modules/Layers.bas)</SwmPath>::<SwmToken path="Modules/Processor.bas" pos="2385:3:3" line-data="        Layers.LoadImageAsNewLayer raiseDialog, processParameters">`LoadImageAsNewLayer`</SwmToken>) --> 519e290ec0b7ab0348806e0f4d64e2af2d9fc7f8f0baef99518dd681638ee3a0(<SwmPath>[Modules/Processor.bas](Modules/Processor.bas)</SwmPath>::Process):::mainFlowStyle
%% 
%% e0eb7a05ee3d732a3e030589917f9ea100a3a88134e1b8f5c0187f7264b31f1f(<SwmPath>[Modules/MoveTool.bas](Modules/MoveTool.bas)</SwmPath>::NotifyKeyDown) --> 519e290ec0b7ab0348806e0f4d64e2af2d9fc7f8f0baef99518dd681638ee3a0(<SwmPath>[Modules/Processor.bas](Modules/Processor.bas)</SwmPath>::Process):::mainFlowStyle
%% 
%% 19b3e55ebbcb2a6666ca9cae9018fcf945a202eb141c649ad360112bde8cce33(<SwmPath>[Modules/SelectionUI.bas](Modules/SelectionUI.bas)</SwmPath>::NotifySelectionMouseUp) --> 519e290ec0b7ab0348806e0f4d64e2af2d9fc7f8f0baef99518dd681638ee3a0(<SwmPath>[Modules/Processor.bas](Modules/Processor.bas)</SwmPath>::Process):::mainFlowStyle
%% 
%% 3c23c9787d44bbfc0a8bcfa6c1ef2949a29b511834fc1ef2cdb87250aebbc28e(<SwmPath>[Forms/File_BatchWizard.frm](Forms/File_BatchWizard.frm)</SwmPath>::cmdNext_Click) --> 9f5e49590f5e5ed922a86696731778583de59723fc9d1ad4ade858bfbe6f9dae(<SwmPath>[Forms/File_BatchWizard.frm](Forms/File_BatchWizard.frm)</SwmPath>::ChangeBatchPage)
%% 
%% 9f5e49590f5e5ed922a86696731778583de59723fc9d1ad4ade858bfbe6f9dae(<SwmPath>[Forms/File_BatchWizard.frm](Forms/File_BatchWizard.frm)</SwmPath>::ChangeBatchPage) --> 0bc4bcdf382f08f7d35863edc83db6d90281bde20daf6b2544d965b2a3ae83f4(<SwmPath>[Forms/File_BatchWizard.frm](Forms/File_BatchWizard.frm)</SwmPath>::PrepareForBatchConversion)
%% 
%% 0bc4bcdf382f08f7d35863edc83db6d90281bde20daf6b2544d965b2a3ae83f4(<SwmPath>[Forms/File_BatchWizard.frm](Forms/File_BatchWizard.frm)</SwmPath>::PrepareForBatchConversion) --> 822dbdc46402979c30729256f4ee31d1b3ae13f2f96487565238e6b45c2c39f2(<SwmPath>[Forms/File_BatchWizard.frm](Forms/File_BatchWizard.frm)</SwmPath>::ApplyEditOperations)
%% 
%% 822dbdc46402979c30729256f4ee31d1b3ae13f2f96487565238e6b45c2c39f2(<SwmPath>[Forms/File_BatchWizard.frm](Forms/File_BatchWizard.frm)</SwmPath>::ApplyEditOperations) --> 519e290ec0b7ab0348806e0f4d64e2af2d9fc7f8f0baef99518dd681638ee3a0(<SwmPath>[Modules/Processor.bas](Modules/Processor.bas)</SwmPath>::Process):::mainFlowStyle
%% 
%% 
%% classDef mainFlowStyle color:#000000,fill:#7CB9F4
%% classDef rootsStyle color:#000000,fill:#00FFF4
%% classDef Style1 color:#000000,fill:#00FFAA
%% classDef Style2 color:#000000,fill:#FFFF00
%% classDef Style3 color:#000000,fill:#AA7CB9
```

# Processing Entry and Pre-Operation Checks

```mermaid
%%{init: {"flowchart": {"defaultRenderer": "elk"}} }%%
flowchart TD
    node1["Start processing user action"]
    click node1 openCode "Modules/Processor.bas:89:142"
    node1 --> node2{"Is this a repeat of last action?"}
    click node2 openCode "Modules/Processor.bas:129:139"
    node2 -->|"Yes"| node3["Restore last action parameters"]
    click node3 openCode "Modules/Processor.bas:130:139"
    node2 -->|"No"| node4["Prepare for processing (UI, parameters, pre-processing)"]
    click node4 openCode "Modules/Processor.bas:142:151"
    node3 --> node4
    node4 --> node5{"Does action require rasterization or selection removal?"}
    click node5 openCode "Modules/Processor.bas:153:172"
    node5 -->|"Yes"| node6["Perform pre-processing (rasterize, remove selection) and exit"]
    click node6 openCode "Modules/Processor.bas:153:161"
    node5 -->|"No"| node7["Notify macro recorder and prepare undo/redo"]
    click node7 openCode "Modules/Processor.bas:174:195"
    node7 --> node8{"Which menu handles the action?"}
    click node8 openCode "Modules/Processor.bas:217:259"
    node8 -->|"File"| node9["File Menu Command Dispatch and Dialog Handling"]
    
    node8 -->|"Edit"| node10["Edit Menu Command Handling"]
    
    node8 -->|"Image"| node11["Image Menu Command Handling and Dialog Routing"]
    
    node8 -->|"Layer"| node12["Layer Operations Dispatch"]
    
    node8 -->|"Select"| node13[Process_SelectMenu]
    click node13 openCode "Modules/Processor.bas:250:253"
    node8 -->|"Adjustments"| node14["Image Adjustments Dispatch"]
    
    node8 -->|"Effects"| node15["Effect Processing and Dialog Routing"]
    
    node8 -->|"Tools"| node16[Process_ToolsMenu]
    click node16 openCode "Modules/Processor.bas:259:320"
    node8 -->|"Unknown"| node17["Prompt user to report unknown action"]
    click node17 openCode "Modules/Processor.bas:323:327"
    node9 --> node18{"Is action 'exit program'?"}
    click node18 openCode "Modules/Processor.bas:224:234"
    node18 -->|"Yes"| node19["Exit PhotoDemon immediately"]
    click node19 openCode "Modules/Processor.bas:226:234"
    node18 -->|"No"| node20["Undo/Redo Finalization and Cancellation Handling"]
    
    node10 --> node20
    node11 --> node20
    node12 --> node20
    node13 --> node20
    node14 --> node20
    node15 --> node20
    node16 --> node20
    node17 --> node20
    node20 --> node21["Update UI and finish"]
    click node21 openCode "Modules/Processor.bas:356:399"
classDef HeadingStyle fill:#777777,stroke:#333,stroke-width:2px;
click node9 goToHeading "File Menu Command Dispatch and Dialog Handling"
node9:::HeadingStyle
click node10 goToHeading "Edit Menu Command Handling"
node10:::HeadingStyle
click node11 goToHeading "Image Menu Command Handling and Dialog Routing"
node11:::HeadingStyle
click node12 goToHeading "Layer Operations Dispatch"
node12:::HeadingStyle
click node14 goToHeading "Image Adjustments Dispatch"
node14:::HeadingStyle
click node15 goToHeading "Effect Processing and Dialog Routing"
node15:::HeadingStyle
click node20 goToHeading "Modules/Processor.bas:166 Finalization and Cancellation Handling"
node20:::HeadingStyle

%% Swimm:
%% %%{init: {"flowchart": {"defaultRenderer": "elk"}} }%%
%% flowchart TD
%%     node1["Start processing user action"]
%%     click node1 openCode "<SwmPath>[Modules/Processor.bas](Modules/Processor.bas)</SwmPath>:89:142"
%%     node1 --> node2{"Is this a repeat of last action?"}
%%     click node2 openCode "<SwmPath>[Modules/Processor.bas](Modules/Processor.bas)</SwmPath>:129:139"
%%     node2 -->|"Yes"| node3["Restore last action parameters"]
%%     click node3 openCode "<SwmPath>[Modules/Processor.bas](Modules/Processor.bas)</SwmPath>:130:139"
%%     node2 -->|"No"| node4["Prepare for processing (UI, parameters, <SwmToken path="Modules/Processor.bas" pos="145:15:17" line-data="    &#39; data types as necessary.  (Some pre-processing steps require parameter knowledge.)">`pre-processing`</SwmToken>)"]
%%     click node4 openCode "<SwmPath>[Modules/Processor.bas](Modules/Processor.bas)</SwmPath>:142:151"
%%     node3 --> node4
%%     node4 --> node5{"Does action require rasterization or selection removal?"}
%%     click node5 openCode "<SwmPath>[Modules/Processor.bas](Modules/Processor.bas)</SwmPath>:153:172"
%%     node5 -->|"Yes"| node6["Perform <SwmToken path="Modules/Processor.bas" pos="145:15:17" line-data="    &#39; data types as necessary.  (Some pre-processing steps require parameter knowledge.)">`pre-processing`</SwmToken> (rasterize, remove selection) and exit"]
%%     click node6 openCode "<SwmPath>[Modules/Processor.bas](Modules/Processor.bas)</SwmPath>:153:161"
%%     node5 -->|"No"| node7["Notify macro recorder and prepare undo/redo"]
%%     click node7 openCode "<SwmPath>[Modules/Processor.bas](Modules/Processor.bas)</SwmPath>:174:195"
%%     node7 --> node8{"Which menu handles the action?"}
%%     click node8 openCode "<SwmPath>[Modules/Processor.bas](Modules/Processor.bas)</SwmPath>:217:259"
%%     node8 -->|"File"| node9["File Menu Command Dispatch and Dialog Handling"]
%%     
%%     node8 -->|"Edit"| node10["Edit Menu Command Handling"]
%%     
%%     node8 -->|"Image"| node11["Image Menu Command Handling and Dialog Routing"]
%%     
%%     node8 -->|"Layer"| node12["Layer Operations Dispatch"]
%%     
%%     node8 -->|"Select"| node13[<SwmToken path="Modules/Processor.bas" pos="250:15:15" line-data="    If (Not processFound) Then processFound = Process_SelectMenu(processID, raiseDialog, processParameters, createUndo, relevantTool, recordAction, returnDetails)">`Process_SelectMenu`</SwmToken>]
%%     click node13 openCode "<SwmPath>[Modules/Processor.bas](Modules/Processor.bas)</SwmPath>:250:253"
%%     node8 -->|"Adjustments"| node14["Image Adjustments Dispatch"]
%%     
%%     node8 -->|"Effects"| node15["Effect Processing and Dialog Routing"]
%%     
%%     node8 -->|"Tools"| node16[<SwmToken path="Modules/Processor.bas" pos="259:15:15" line-data="    If (Not processFound) Then processFound = Process_ToolsMenu(processID, raiseDialog, processParameters, createUndo, relevantTool, recordAction, returnDetails)">`Process_ToolsMenu`</SwmToken>]
%%     click node16 openCode "<SwmPath>[Modules/Processor.bas](Modules/Processor.bas)</SwmPath>:259:320"
%%     node8 -->|"Unknown"| node17["Prompt user to report unknown action"]
%%     click node17 openCode "<SwmPath>[Modules/Processor.bas](Modules/Processor.bas)</SwmPath>:323:327"
%%     node9 --> node18{"Is action 'exit program'?"}
%%     click node18 openCode "<SwmPath>[Modules/Processor.bas](Modules/Processor.bas)</SwmPath>:224:234"
%%     node18 -->|"Yes"| node19["Exit <SwmToken path="Modules/Processor.bas" pos="206:11:11" line-data="    &#39; list of every possible PhotoDemon action, filter, or other operation.  Depending on the processID, additional">`PhotoDemon`</SwmToken> immediately"]
%%     click node19 openCode "<SwmPath>[Modules/Processor.bas](Modules/Processor.bas)</SwmPath>:226:234"
%%     node18 -->|"No"| node20["<SwmToken path="Modules/Processor.bas" pos="166:33:35" line-data="    &#39; if selections persist after a size change - and this is particularly relevant for the Undo/Redo engine,">`Undo/Redo`</SwmToken> Finalization and Cancellation Handling"]
%%     
%%     node10 --> node20
%%     node11 --> node20
%%     node12 --> node20
%%     node13 --> node20
%%     node14 --> node20
%%     node15 --> node20
%%     node16 --> node20
%%     node17 --> node20
%%     node20 --> node21["Update UI and finish"]
%%     click node21 openCode "<SwmPath>[Modules/Processor.bas](Modules/Processor.bas)</SwmPath>:356:399"
%% classDef HeadingStyle fill:#777777,stroke:#333,stroke-width:2px;
%% click node9 goToHeading "File Menu Command Dispatch and Dialog Handling"
%% node9:::HeadingStyle
%% click node10 goToHeading "Edit Menu Command Handling"
%% node10:::HeadingStyle
%% click node11 goToHeading "Image Menu Command Handling and Dialog Routing"
%% node11:::HeadingStyle
%% click node12 goToHeading "Layer Operations Dispatch"
%% node12:::HeadingStyle
%% click node14 goToHeading "Image Adjustments Dispatch"
%% node14:::HeadingStyle
%% click node15 goToHeading "Effect Processing and Dialog Routing"
%% node15:::HeadingStyle
%% click node20 goToHeading "<SwmToken path="Modules/Processor.bas" pos="166:33:35" line-data="    &#39; if selections persist after a size change - and this is particularly relevant for the Undo/Redo engine,">`Undo/Redo`</SwmToken> Finalization and Cancellation Handling"
%% node20:::HeadingStyle
```

<SwmSnippet path="/Modules/Processor.bas" line="89">

---

In <SwmToken path="Modules/Processor.bas" pos="89:4:4" line-data="Public Sub Process(ByVal processID As String, Optional raiseDialog As Boolean = False, Optional processParameters As String = vbNullString, Optional createUndo As PD_UndoType = UNDO_Nothing, Optional relevantTool As Long = -1, Optional recordAction As Boolean = True)">`Process`</SwmToken>, we prep for an operation by logging, parameter handling, and making sure the UI is locked down (<SwmToken path="Modules/Processor.bas" pos="142:1:1" line-data="    SetProcessorUI_Busy processID, raiseDialog, processParameters, createUndo, relevantTool, recordAction">`SetProcessorUI_Busy`</SwmToken>) before any real work starts. This keeps the user from messing with things mid-operation and avoids weird state bugs.

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

<SwmToken path="Modules/Processor.bas" pos="1256:4:4" line-data="Private Sub SetProcessorUI_Busy(ByVal processID As String, Optional raiseDialog As Boolean = False, Optional processParameters As String = vbNullString, Optional createUndo As PD_UndoType = UNDO_Nothing, Optional relevantTool As Long = -1, Optional recordAction As Boolean = True)">`SetProcessorUI_Busy`</SwmToken> decides if the UI should look busy (hourglass, etc), marks the processor as running, and updates the UI state, but only for real processing, not just dialogs.

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

After <SwmToken path="Modules/Processor.bas" pos="142:1:1" line-data="    SetProcessorUI_Busy processID, raiseDialog, processParameters, createUndo, relevantTool, recordAction">`SetProcessorUI_Busy`</SwmToken>, <SwmToken path="Modules/Processor.bas" pos="89:4:4" line-data="Public Sub Process(ByVal processID As String, Optional raiseDialog As Boolean = False, Optional processParameters As String = vbNullString, Optional createUndo As PD_UndoType = UNDO_Nothing, Optional relevantTool As Long = -1, Optional recordAction As Boolean = True)">`Process`</SwmToken> sets up parameter parsing and calls <SwmToken path="Modules/Processor.bas" pos="151:1:1" line-data="    Processor_BeforeStarting processID, raiseDialog, processParameters, createUndo, relevantTool, recordAction">`Processor_BeforeStarting`</SwmToken> to handle any weird <SwmToken path="Modules/Processor.bas" pos="145:15:17" line-data="    &#39; data types as necessary.  (Some pre-processing steps require parameter knowledge.)">`pre-processing`</SwmToken> needs, especially for commands that mess with selections.

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

<SwmToken path="Modules/Processor.bas" pos="705:4:4" line-data="Public Sub Processor_BeforeStarting(ByVal processID As String, Optional raiseDialog As Boolean = False, Optional processParameters As String = vbNullString, Optional createUndo As PD_UndoType = UNDO_Nothing, Optional relevantTool As Long = -1, Optional recordAction As Boolean = True)">`Processor_BeforeStarting`</SwmToken> checks if we're about to Crop (with no dialog and not in batch mode). If so, it tells the undo manager to save both image and selection state before the crop, so undo/redo won't break when the selection is cleared.

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

After <SwmToken path="Modules/Processor.bas" pos="151:1:1" line-data="    Processor_BeforeStarting processID, raiseDialog, processParameters, createUndo, relevantTool, recordAction">`Processor_BeforeStarting`</SwmToken>, <SwmToken path="Modules/Processor.bas" pos="89:4:4" line-data="Public Sub Process(ByVal processID As String, Optional raiseDialog As Boolean = False, Optional processParameters As String = vbNullString, Optional createUndo As PD_UndoType = UNDO_Nothing, Optional relevantTool As Long = -1, Optional recordAction As Boolean = True)">`Process`</SwmToken> checks if vector layers need rasterizing before continuing. If the user cancels, we exit out immediately.

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

<SwmToken path="Modules/Processor.bas" pos="1107:4:4" line-data="Private Function CheckRasterizeRequirements(ByVal processID As String, Optional raiseDialog As Boolean = False, Optional processParameters As String = vbNullString, Optional createUndo As PD_UndoType = UNDO_Nothing) As Boolean">`CheckRasterizeRequirements`</SwmToken> checks if the current operation will mess with vector layers. For <SwmToken path="Modules/Processor.bas" pos="1187:19:21" line-data="        &#39;If we need to do a &quot;special-case whole-image&quot; rasterization, do so now.">`whole-image`</SwmToken> ops, it looks for vector layers and applies exceptions (like merge/crop). If rasterization is needed, it prompts the user and rasterizes the right layers if they agree. For single-layer ops, it checks if the active layer is vector and does the same. If the user cancels, we return False to abort the operation.

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

After <SwmToken path="Modules/Processor.bas" pos="158:6:6" line-data="    If (Not CheckRasterizeRequirements(processID, raiseDialog, processParameters, createUndo)) Then">`CheckRasterizeRequirements`</SwmToken>, if the user cancels, we call <SwmToken path="Modules/Processor.bas" pos="159:1:1" line-data="        SetProcessorUI_Idle processID, raiseDialog, processParameters, createUndo, relevantTool, recordAction">`SetProcessorUI_Idle`</SwmToken> to clean up and exit.

```visual basic
        SetProcessorUI_Idle processID, raiseDialog, processParameters, createUndo, relevantTool, recordAction
        Exit Sub
    End If
    
```

---

</SwmSnippet>

<SwmSnippet path="/Modules/Processor.bas" line="1275">

---

<SwmToken path="Modules/Processor.bas" pos="1275:4:4" line-data="Private Sub SetProcessorUI_Idle(ByVal processID As String, Optional raiseDialog As Boolean = False, Optional processParameters As String = vbNullString, Optional createUndo As PD_UndoType = UNDO_Nothing, Optional relevantTool As Long = -1, Optional recordAction As Boolean = True)">`SetProcessorUI_Idle`</SwmToken> unlocks the UI, decrements the nested count, and restores focus if we're done.

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

Back in <SwmToken path="Modules/Processor.bas" pos="89:4:4" line-data="Public Sub Process(ByVal processID As String, Optional raiseDialog As Boolean = False, Optional processParameters As String = vbNullString, Optional createUndo As PD_UndoType = UNDO_Nothing, Optional relevantTool As Long = -1, Optional recordAction As Boolean = True)">`Process`</SwmToken>, after cleaning up from a cancelled rasterization, we check if the current operation requires removing the selection (like resizing or rotating). If so, we trigger selection removal to keep the selection mask and image in sync and avoid undo/redo issues.

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

<SwmToken path="Modules/Processor.bas" pos="1059:4:4" line-data="Private Function RemoveSelectionAsNecessary(ByVal processID As String, Optional raiseDialog As Boolean = False, Optional processParameters As String = vbNullString, Optional createUndo As PD_UndoType = UNDO_Nothing) As Boolean">`RemoveSelectionAsNecessary`</SwmToken> checks if there's an active selection and if the current operation is one that will mess with the selection mask (resize, rotate, flip, etc.). If so, it calls <SwmToken path="Modules/Processor.bas" pos="1094:1:3" line-data="                Processor.Process &quot;Remove selection&quot;, , , UNDO_Selection">`Processor.Process`</SwmToken> to remove the selection before continuing.

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

Back in <SwmToken path="Modules/Processor.bas" pos="89:4:4" line-data="Public Sub Process(ByVal processID As String, Optional raiseDialog As Boolean = False, Optional processParameters As String = vbNullString, Optional createUndo As PD_UndoType = UNDO_Nothing, Optional relevantTool As Long = -1, Optional recordAction As Boolean = True)">`Process`</SwmToken>, after handling selection removal, we notify the macro recorder, update undo state if needed, and call <SwmToken path="Modules/Processor.bas" pos="193:1:1" line-data="        CheckForCanvasModifications createUndo">`CheckForCanvasModifications`</SwmToken> to make sure any unsaved selection changes are captured before the main operation runs.

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

<SwmToken path="Modules/Processor.bas" pos="1006:4:4" line-data="Private Sub CheckForCanvasModifications(ByVal createUndo As PD_UndoType)">`CheckForCanvasModifications`</SwmToken> checks if there's an active selection and if the undo type isn't already about selection or everything. It compares the last saved selection XML with the current one, and if they're different, it creates a new undo entry for the selection change.

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

Back in <SwmToken path="Modules/Processor.bas" pos="89:4:4" line-data="Public Sub Process(ByVal processID As String, Optional raiseDialog As Boolean = False, Optional processParameters As String = vbNullString, Optional createUndo As PD_UndoType = UNDO_Nothing, Optional relevantTool As Long = -1, Optional recordAction As Boolean = True)">`Process`</SwmToken>, after all the pre-checks and undo handling, we hit the main dispatch section. Here, we check the <SwmToken path="Modules/Processor.bas" pos="205:28:28" line-data="    &#39;The bulk of this routine starts here.  From this point on, the processID string is compared against a hard-coded">`processID`</SwmToken> against all known menu operations (File, Edit, Image, etc.) and call the right handler. If it's a File menu action, we call <SwmToken path="Modules/Processor.bas" pos="217:5:5" line-data="    processFound = Process_FileMenu(processID, raiseDialog, processParameters, createUndo, relevantTool, recordAction, returnDetails)">`Process_FileMenu`</SwmToken> next.

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
    click node3 openCode "Modules/Processor.bas:1361:1361"
    node3 -->|"Yes"| node4["Show new image dialog"]
    click node4 openCode "Modules/Interface.bas:993:1095"
    node3 -->|"No"| node5["Create new image"]
    click node5 openCode "Modules/Processor.bas:1361:1361"
    node2 -->|"Open"| node6["Open image"]
    click node6 openCode "Modules/Processor.bas:1365:1365"
    node2 -->|"Close/Close all"| node7["Close image(s)"]
    click node7 openCode "Modules/Processor.bas:1369:1374"
    node2 -->|"Save/Save as/Save copy"| node8["Save image (various modes)"]
    click node8 openCode "Modules/Processor.bas:1377:1386"
    node2 -->|"Revert"| node9{"Is revert enabled?"}
    click node9 openCode "Modules/Processor.bas:1389:1392"
    node9 -->|"Yes"| node10["Revert image"]
    click node10 openCode "Modules/Processor.bas:1390:1391"
    node9 -->|"No"| node11["Skip revert"]
    click node11 openCode "Modules/Processor.bas:1392:1392"
    node2 -->|"Export (image/animation/color lookup/profile/palette)"| node12["Export image data"]
    click node12 openCode "Modules/Processor.bas:1396:1423"
    node2 -->|"Export layers"| node13{"Is export layers enabled?"}
    click node13 openCode "Modules/Processor.bas:1400:1406"
    node13 -->|"Yes"| node14{"Show dialog?"}
    click node14 openCode "Modules/Processor.bas:1401:1403"
    node14 -->|"Yes"| node15["Show export layers dialog"]
    click node15 openCode "Modules/Interface.bas:993:1095"
    node14 -->|"No"| node16["Export layers"]
    click node16 openCode "Modules/Processor.bas:1403:1403"
    node13 -->|"No"| node17["Skip export layers"]
    click node17 openCode "Modules/Processor.bas:1406:1406"
    node2 -->|"Batch wizard"| node18["Show batch wizard dialog"]
    click node18 openCode "Modules/Processor.bas:1426:1426"
    node2 -->|"Print"| node19{"Show dialog?"}
    click node19 openCode "Modules/Processor.bas:1430:1440"
    node19 -->|"Yes"| node20{"Is OS Vista or later?"}
    click node20 openCode "Modules/Processor.bas:1434:1438"
    node20 -->|"Yes"| node21["Print via Windows photo printer"]
    click node21 openCode "Modules/Processor.bas:1435:1435"
    node20 -->|"No"| node22["Show legacy print dialog"]
    click node22 openCode "Modules/Processor.bas:1437:1438"
    node19 -->|"No"| node23["Skip print"]
    click node23 openCode "Modules/Processor.bas:1440:1440"
    node2 -->|"Exit program"| node24["Trigger program exit"]
    click node24 openCode "Modules/Processor.bas:1446:1447"
    node2 -->|"Select scanner/camera"| node25["Select scanner/camera"]
    click node25 openCode "Modules/Processor.bas:1450:1450"
    node2 -->|"Scan image"| node26["Scan image"]
    click node26 openCode "Modules/Processor.bas:1454:1454"
    node2 -->|"Screen capture"| node27{"Show dialog?"}
    click node27 openCode "Modules/Processor.bas:1458:1458"
    node27 -->|"Yes"| node28["Show screen capture dialog"]
    click node28 openCode "Modules/Interface.bas:993:1095"
    node27 -->|"No"| node29["Capture screen"]
    click node29 openCode "Modules/Processor.bas:1458:1458"
    node2 -->|"Internet import"| node30{"Show dialog?"}
    click node30 openCode "Modules/Processor.bas:1462:1462"
    node30 -->|"Yes"| node31["Show internet import dialog"]
    click node31 openCode "Modules/Interface.bas:993:1095"
    node30 -->|"No"| node32["Skip internet import"]
    click node32 openCode "Modules/Processor.bas:1463:1463"

classDef HeadingStyle fill:#777777,stroke:#333,stroke-width:2px;

%% Swimm:
%% %%{init: {"flowchart": {"defaultRenderer": "elk"}} }%%
%% flowchart TD
%%     node1["User selects a File menu command"]
%%     click node1 openCode "<SwmPath>[Modules/Processor.bas](Modules/Processor.bas)</SwmPath>:1358:1467"
%%     node1 --> node2{"Which command?"}
%%     click node2 openCode "<SwmPath>[Modules/Processor.bas](Modules/Processor.bas)</SwmPath>:1360:1465"
%%     node2 -->|"New image"| node3{"Show dialog?"}
%%     click node3 openCode "<SwmPath>[Modules/Processor.bas](Modules/Processor.bas)</SwmPath>:1361:1361"
%%     node3 -->|"Yes"| node4["Show new image dialog"]
%%     click node4 openCode "<SwmPath>[Modules/Interface.bas](Modules/Interface.bas)</SwmPath>:993:1095"
%%     node3 -->|"No"| node5["Create new image"]
%%     click node5 openCode "<SwmPath>[Modules/Processor.bas](Modules/Processor.bas)</SwmPath>:1361:1361"
%%     node2 -->|"Open"| node6["Open image"]
%%     click node6 openCode "<SwmPath>[Modules/Processor.bas](Modules/Processor.bas)</SwmPath>:1365:1365"
%%     node2 -->|"Close/Close all"| node7["Close image(s)"]
%%     click node7 openCode "<SwmPath>[Modules/Processor.bas](Modules/Processor.bas)</SwmPath>:1369:1374"
%%     node2 -->|"Save/Save as/Save copy"| node8["Save image (various modes)"]
%%     click node8 openCode "<SwmPath>[Modules/Processor.bas](Modules/Processor.bas)</SwmPath>:1377:1386"
%%     node2 -->|"Revert"| node9{"Is revert enabled?"}
%%     click node9 openCode "<SwmPath>[Modules/Processor.bas](Modules/Processor.bas)</SwmPath>:1389:1392"
%%     node9 -->|"Yes"| node10["Revert image"]
%%     click node10 openCode "<SwmPath>[Modules/Processor.bas](Modules/Processor.bas)</SwmPath>:1390:1391"
%%     node9 -->|"No"| node11["Skip revert"]
%%     click node11 openCode "<SwmPath>[Modules/Processor.bas](Modules/Processor.bas)</SwmPath>:1392:1392"
%%     node2 -->|"Export (image/animation/color lookup/profile/palette)"| node12["Export image data"]
%%     click node12 openCode "<SwmPath>[Modules/Processor.bas](Modules/Processor.bas)</SwmPath>:1396:1423"
%%     node2 -->|"Export layers"| node13{"Is export layers enabled?"}
%%     click node13 openCode "<SwmPath>[Modules/Processor.bas](Modules/Processor.bas)</SwmPath>:1400:1406"
%%     node13 -->|"Yes"| node14{"Show dialog?"}
%%     click node14 openCode "<SwmPath>[Modules/Processor.bas](Modules/Processor.bas)</SwmPath>:1401:1403"
%%     node14 -->|"Yes"| node15["Show export layers dialog"]
%%     click node15 openCode "<SwmPath>[Modules/Interface.bas](Modules/Interface.bas)</SwmPath>:993:1095"
%%     node14 -->|"No"| node16["Export layers"]
%%     click node16 openCode "<SwmPath>[Modules/Processor.bas](Modules/Processor.bas)</SwmPath>:1403:1403"
%%     node13 -->|"No"| node17["Skip export layers"]
%%     click node17 openCode "<SwmPath>[Modules/Processor.bas](Modules/Processor.bas)</SwmPath>:1406:1406"
%%     node2 -->|"Batch wizard"| node18["Show batch wizard dialog"]
%%     click node18 openCode "<SwmPath>[Modules/Processor.bas](Modules/Processor.bas)</SwmPath>:1426:1426"
%%     node2 -->|"Print"| node19{"Show dialog?"}
%%     click node19 openCode "<SwmPath>[Modules/Processor.bas](Modules/Processor.bas)</SwmPath>:1430:1440"
%%     node19 -->|"Yes"| node20{"Is OS Vista or later?"}
%%     click node20 openCode "<SwmPath>[Modules/Processor.bas](Modules/Processor.bas)</SwmPath>:1434:1438"
%%     node20 -->|"Yes"| node21["Print via Windows photo printer"]
%%     click node21 openCode "<SwmPath>[Modules/Processor.bas](Modules/Processor.bas)</SwmPath>:1435:1435"
%%     node20 -->|"No"| node22["Show legacy print dialog"]
%%     click node22 openCode "<SwmPath>[Modules/Processor.bas](Modules/Processor.bas)</SwmPath>:1437:1438"
%%     node19 -->|"No"| node23["Skip print"]
%%     click node23 openCode "<SwmPath>[Modules/Processor.bas](Modules/Processor.bas)</SwmPath>:1440:1440"
%%     node2 -->|"Exit program"| node24["Trigger program exit"]
%%     click node24 openCode "<SwmPath>[Modules/Processor.bas](Modules/Processor.bas)</SwmPath>:1446:1447"
%%     node2 -->|"Select scanner/camera"| node25["Select scanner/camera"]
%%     click node25 openCode "<SwmPath>[Modules/Processor.bas](Modules/Processor.bas)</SwmPath>:1450:1450"
%%     node2 -->|"Scan image"| node26["Scan image"]
%%     click node26 openCode "<SwmPath>[Modules/Processor.bas](Modules/Processor.bas)</SwmPath>:1454:1454"
%%     node2 -->|"Screen capture"| node27{"Show dialog?"}
%%     click node27 openCode "<SwmPath>[Modules/Processor.bas](Modules/Processor.bas)</SwmPath>:1458:1458"
%%     node27 -->|"Yes"| node28["Show screen capture dialog"]
%%     click node28 openCode "<SwmPath>[Modules/Interface.bas](Modules/Interface.bas)</SwmPath>:993:1095"
%%     node27 -->|"No"| node29["Capture screen"]
%%     click node29 openCode "<SwmPath>[Modules/Processor.bas](Modules/Processor.bas)</SwmPath>:1458:1458"
%%     node2 -->|"Internet import"| node30{"Show dialog?"}
%%     click node30 openCode "<SwmPath>[Modules/Processor.bas](Modules/Processor.bas)</SwmPath>:1462:1462"
%%     node30 -->|"Yes"| node31["Show internet import dialog"]
%%     click node31 openCode "<SwmPath>[Modules/Interface.bas](Modules/Interface.bas)</SwmPath>:993:1095"
%%     node30 -->|"No"| node32["Skip internet import"]
%%     click node32 openCode "<SwmPath>[Modules/Processor.bas](Modules/Processor.bas)</SwmPath>:1463:1463"
%% 
%% classDef HeadingStyle fill:#777777,stroke:#333,stroke-width:2px;
```

<SwmSnippet path="/Modules/Processor.bas" line="1358">

---

In <SwmToken path="Modules/Processor.bas" pos="1358:4:4" line-data="Private Function Process_FileMenu(ByVal processID As String, Optional raiseDialog As Boolean = False, Optional processParameters As String = vbNullString, Optional createUndo As PD_UndoType = UNDO_Nothing, Optional relevantTool As Long = -1, Optional recordAction As Boolean = True, Optional ByRef returnDetails As String = vbNullString) As Boolean">`Process_FileMenu`</SwmToken>, we match the <SwmToken path="Modules/Processor.bas" pos="1358:8:8" line-data="Private Function Process_FileMenu(ByVal processID As String, Optional raiseDialog As Boolean = False, Optional processParameters As String = vbNullString, Optional createUndo As PD_UndoType = UNDO_Nothing, Optional relevantTool As Long = -1, Optional recordAction As Boolean = True, Optional ByRef returnDetails As String = vbNullString) As Boolean">`processID`</SwmToken> against known file commands. For dialog-based actions (like 'New image'), if <SwmToken path="Modules/Processor.bas" pos="1358:17:17" line-data="Private Function Process_FileMenu(ByVal processID As String, Optional raiseDialog As Boolean = False, Optional processParameters As String = vbNullString, Optional createUndo As PD_UndoType = UNDO_Nothing, Optional relevantTool As Long = -1, Optional recordAction As Boolean = True, Optional ByRef returnDetails As String = vbNullString) As Boolean">`raiseDialog`</SwmToken> is set, we show the dialog (handled by <SwmToken path="Modules/Processor.bas" pos="1361:7:7" line-data="        If raiseDialog Then ShowPDDialog vbModal, FormNewImage Else FileMenu.CreateNewImage processParameters">`ShowPDDialog`</SwmToken> in <SwmPath>[Modules/Interface.bas](Modules/Interface.bas)</SwmPath>); otherwise, we run the action directly. This pattern repeats for each file menu command.

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

<SwmToken path="Modules/Interface.bas" pos="993:4:4" line-data="Public Sub ShowPDDialog(ByRef dialogModality As FormShowConstants, ByRef dialogForm As Form, Optional ByVal doNotUnload As Boolean = False)">`ShowPDDialog`</SwmToken> handles dialog stacking by assigning ownership based on whether another modal dialog is already open. It centers the dialog on the main form unless a previous position is stored, and makes sure the dialog doesn't go <SwmToken path="Modules/Interface.bas" pos="1057:18:20" line-data="        &#39;If this position results in the dialog sitting off-screen, move it so that its bottom-right corner is always on-screen.">`off-screen`</SwmToken> (especially the <SwmToken path="Modules/Interface.bas" pos="1057:33:35" line-data="        &#39;If this position results in the dialog sitting off-screen, move it so that its bottom-right corner is always on-screen.">`bottom-right`</SwmToken> corner). It also mirrors window icons for proper Alt+Tab behavior, then restores them after the dialog closes.

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

After returning from <SwmToken path="Modules/Processor.bas" pos="1426:3:3" line-data="        Interface.ShowPDDialog vbModal, FormBatchWizard">`ShowPDDialog`</SwmToken> (or running the direct action), <SwmToken path="Modules/Processor.bas" pos="1407:1:1" line-data="        Process_FileMenu = True">`Process_FileMenu`</SwmToken> continues matching <SwmToken path="Modules/Processor.bas" pos="209:15:15" line-data="    &#39;For ease of reference, the various processIDs are divided into categories of similar functions.  These categories">`processIDs`</SwmToken> for other file commands. It handles OS-specific print dialogs, signals the main process to exit if needed, and calls out to plugins for scanner actions. Each recognized command sets <SwmToken path="Modules/Processor.bas" pos="1407:1:1" line-data="        Process_FileMenu = True">`Process_FileMenu`</SwmToken> = True to indicate it was handled.

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

## Edit Menu Dispatch and Post-File Operations

```mermaid
%%{init: {"flowchart": {"defaultRenderer": "elk"}} }%%
flowchart TD
    node1{"Was a menu operation found?"}
    click node1 openCode "Modules/Processor.bas:221:238"
    node1 -->|"Yes"| node2{"Is returnDetails 'exit program'?"}
    click node2 openCode "Modules/Processor.bas:224:236"
    node1 -->|"No"| node5["Process edit menu operation"]
    click node5 openCode "Modules/Processor.bas:241:241"
    node2 -->|"Yes"| node6["Unload main window"]
    click node6 openCode "Modules/Processor.bas:226:227"
    node2 -->|"No"| node5
    node6 --> node3{"Is program shutting down?"}
    click node3 openCode "Modules/Processor.bas:231:234"
    node3 -->|"Yes"| node4["Stop all further processing"]
    click node4 openCode "Modules/Processor.bas:232:233"
    node3 -->|"No"| node5
classDef HeadingStyle fill:#777777,stroke:#333,stroke-width:2px;

%% Swimm:
%% %%{init: {"flowchart": {"defaultRenderer": "elk"}} }%%
%% flowchart TD
%%     node1{"Was a menu operation found?"}
%%     click node1 openCode "<SwmPath>[Modules/Processor.bas](Modules/Processor.bas)</SwmPath>:221:238"
%%     node1 -->|"Yes"| node2{"Is <SwmToken path="Modules/Processor.bas" pos="216:10:10" line-data="    Dim processFound As Boolean, returnDetails As String">`returnDetails`</SwmToken> 'exit program'?"}
%%     click node2 openCode "<SwmPath>[Modules/Processor.bas](Modules/Processor.bas)</SwmPath>:224:236"
%%     node1 -->|"No"| node5["Process edit menu operation"]
%%     click node5 openCode "<SwmPath>[Modules/Processor.bas](Modules/Processor.bas)</SwmPath>:241:241"
%%     node2 -->|"Yes"| node6["Unload main window"]
%%     click node6 openCode "<SwmPath>[Modules/Processor.bas](Modules/Processor.bas)</SwmPath>:226:227"
%%     node2 -->|"No"| node5
%%     node6 --> node3{"Is program shutting down?"}
%%     click node3 openCode "<SwmPath>[Modules/Processor.bas](Modules/Processor.bas)</SwmPath>:231:234"
%%     node3 -->|"Yes"| node4["Stop all further processing"]
%%     click node4 openCode "<SwmPath>[Modules/Processor.bas](Modules/Processor.bas)</SwmPath>:232:233"
%%     node3 -->|"No"| node5
%% classDef HeadingStyle fill:#777777,stroke:#333,stroke-width:2px;
```

<SwmSnippet path="/Modules/Processor.bas" line="219">

---

Back in <SwmToken path="Modules/Processor.bas" pos="89:4:4" line-data="Public Sub Process(ByVal processID As String, Optional raiseDialog As Boolean = False, Optional processParameters As String = vbNullString, Optional createUndo As PD_UndoType = UNDO_Nothing, Optional relevantTool As Long = -1, Optional recordAction As Boolean = True)">`Process`</SwmToken>, after handling file menu actions (and any special exit codes), we check if the <SwmToken path="Modules/Processor.bas" pos="241:17:17" line-data="    If (Not processFound) Then processFound = Process_EditMenu(processID, raiseDialog, processParameters, createUndo, relevantTool, recordAction, returnDetails)">`processID`</SwmToken> matches any edit menu actions. If so, we dispatch to <SwmToken path="Modules/Processor.bas" pos="241:15:15" line-data="    If (Not processFound) Then processFound = Process_EditMenu(processID, raiseDialog, processParameters, createUndo, relevantTool, recordAction, returnDetails)">`Process_EditMenu`</SwmToken> next.

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

## Edit Menu Command Handling

```mermaid
%%{init: {"flowchart": {"defaultRenderer": "elk"}} }%%
flowchart TD
    node1["User selects an Edit menu action"] --> node2{"Which action is requested?"}
    click node1 openCode "Modules/Processor.bas:1472:1604"
    node2 -->|"Undo/Redo/Undo history"| node3["Perform Undo, Redo, or Undo history"]
    click node2 openCode "Modules/Processor.bas:1478:1513"
    node2 -->|"Fade"| node4{"Show dialog?"}
    click node2 openCode "Modules/Processor.bas:1514:1517"
    node4 -->|"Yes"| node5["Show Fade dialog"]
    click node5 openCode "Modules/Processor.bas:1515:1516"
    node4 -->|"No"| node6["Perform Fade operation"]
    click node6 openCode "Modules/Processor.bas:1515:1516"
    node2 -->|"Clipboard (Cut, Copy, Paste, etc.)"| node7{"Is special clipboard action?"}
    click node2 openCode "Modules/Processor.bas:1518:1586"
    node7 -->|"Yes"| node8{"Show dialog?"}
    click node7 openCode "Modules/Processor.bas:1560:1574"
    node8 -->|"Yes"| node9["Show clipboard dialog"]
    click node9 openCode "Modules/Processor.bas:1562:1570"
    node8 -->|"No"| node10["Perform special clipboard operation"]
    click node10 openCode "Modules/Processor.bas:1564:1572"
    node7 -->|"No"| node11{"Is image active for Paste?"}
    click node11 openCode "Modules/Processor.bas:1538:1554"
    node11 -->|"Yes"| node12["Perform clipboard operation"]
    click node12 openCode "Modules/Processor.bas:1544:1548"
    node11 -->|"No"| node13["Paste as new image"]
    click node13 openCode "Modules/Processor.bas:1557:1558"
    node2 -->|"Selection (Clear, Fill, Stroke, etc.)"| node14["Perform selection operation"]
    click node14 openCode "Modules/Processor.bas:1588:1602"
    node3 --> node15{"Was Undo/Redo used?"}
    click node15 openCode "Modules/Processor.bas:1606:1612"
    node15 -->|"Yes"| node16["Sync layer settings"]
    click node16 openCode "Modules/Processor.bas:1609:1610"
    node15 -->|"No"| node17["End"]
    node5 --> node17
    node6 --> node17
    node9 --> node17
    node10 --> node17
    node12 --> node17
    node13 --> node17
    node14 --> node17
    node16 --> node17
classDef HeadingStyle fill:#777777,stroke:#333,stroke-width:2px;

%% Swimm:
%% %%{init: {"flowchart": {"defaultRenderer": "elk"}} }%%
%% flowchart TD
%%     node1["User selects an Edit menu action"] --> node2{"Which action is requested?"}
%%     click node1 openCode "<SwmPath>[Modules/Processor.bas](Modules/Processor.bas)</SwmPath>:1472:1604"
%%     node2 -->|"Undo/Redo/Undo history"| node3["Perform Undo, Redo, or Undo history"]
%%     click node2 openCode "<SwmPath>[Modules/Processor.bas](Modules/Processor.bas)</SwmPath>:1478:1513"
%%     node2 -->|"Fade"| node4{"Show dialog?"}
%%     click node2 openCode "<SwmPath>[Modules/Processor.bas](Modules/Processor.bas)</SwmPath>:1514:1517"
%%     node4 -->|"Yes"| node5["Show Fade dialog"]
%%     click node5 openCode "<SwmPath>[Modules/Processor.bas](Modules/Processor.bas)</SwmPath>:1515:1516"
%%     node4 -->|"No"| node6["Perform Fade operation"]
%%     click node6 openCode "<SwmPath>[Modules/Processor.bas](Modules/Processor.bas)</SwmPath>:1515:1516"
%%     node2 -->|"Clipboard (Cut, Copy, Paste, etc.)"| node7{"Is special clipboard action?"}
%%     click node2 openCode "<SwmPath>[Modules/Processor.bas](Modules/Processor.bas)</SwmPath>:1518:1586"
%%     node7 -->|"Yes"| node8{"Show dialog?"}
%%     click node7 openCode "<SwmPath>[Modules/Processor.bas](Modules/Processor.bas)</SwmPath>:1560:1574"
%%     node8 -->|"Yes"| node9["Show clipboard dialog"]
%%     click node9 openCode "<SwmPath>[Modules/Processor.bas](Modules/Processor.bas)</SwmPath>:1562:1570"
%%     node8 -->|"No"| node10["Perform special clipboard operation"]
%%     click node10 openCode "<SwmPath>[Modules/Processor.bas](Modules/Processor.bas)</SwmPath>:1564:1572"
%%     node7 -->|"No"| node11{"Is image active for Paste?"}
%%     click node11 openCode "<SwmPath>[Modules/Processor.bas](Modules/Processor.bas)</SwmPath>:1538:1554"
%%     node11 -->|"Yes"| node12["Perform clipboard operation"]
%%     click node12 openCode "<SwmPath>[Modules/Processor.bas](Modules/Processor.bas)</SwmPath>:1544:1548"
%%     node11 -->|"No"| node13["Paste as new image"]
%%     click node13 openCode "<SwmPath>[Modules/Processor.bas](Modules/Processor.bas)</SwmPath>:1557:1558"
%%     node2 -->|"Selection (Clear, Fill, Stroke, etc.)"| node14["Perform selection operation"]
%%     click node14 openCode "<SwmPath>[Modules/Processor.bas](Modules/Processor.bas)</SwmPath>:1588:1602"
%%     node3 --> node15{"Was <SwmToken path="Modules/Processor.bas" pos="166:33:35" line-data="    &#39; if selections persist after a size change - and this is particularly relevant for the Undo/Redo engine,">`Undo/Redo`</SwmToken> used?"}
%%     click node15 openCode "<SwmPath>[Modules/Processor.bas](Modules/Processor.bas)</SwmPath>:1606:1612"
%%     node15 -->|"Yes"| node16["Sync layer settings"]
%%     click node16 openCode "<SwmPath>[Modules/Processor.bas](Modules/Processor.bas)</SwmPath>:1609:1610"
%%     node15 -->|"No"| node17["End"]
%%     node5 --> node17
%%     node6 --> node17
%%     node9 --> node17
%%     node10 --> node17
%%     node12 --> node17
%%     node13 --> node17
%%     node14 --> node17
%%     node16 --> node17
%% classDef HeadingStyle fill:#777777,stroke:#333,stroke-width:2px;
```

<SwmSnippet path="/Modules/Processor.bas" line="1472">

---

In <SwmToken path="Modules/Processor.bas" pos="1472:4:4" line-data="Private Function Process_EditMenu(ByRef processID As String, Optional ByVal raiseDialog As Boolean = False, Optional ByRef processParameters As String = vbNullString, Optional ByRef createUndo As PD_UndoType = UNDO_Nothing, Optional ByRef relevantTool As Long = -1, Optional ByRef recordAction As Boolean = True, Optional ByRef returnDetails As String = vbNullString) As Boolean">`Process_EditMenu`</SwmToken>, we check if the <SwmToken path="Modules/Processor.bas" pos="1472:8:8" line-data="Private Function Process_EditMenu(ByRef processID As String, Optional ByVal raiseDialog As Boolean = False, Optional ByRef processParameters As String = vbNullString, Optional ByRef createUndo As PD_UndoType = UNDO_Nothing, Optional ByRef relevantTool As Long = -1, Optional ByRef recordAction As Boolean = True, Optional ByRef returnDetails As String = vbNullString) As Boolean">`processID`</SwmToken> is undo/redo/history and handle those by restoring the right state, then notifying the interface and viewport to update UI elements. Clipboard actions are routed through <SwmToken path="Modules/Processor.bas" pos="1519:1:1" line-data="        g_Clipboard.ClipboardCut False">`g_Clipboard`</SwmToken>, and for 'Paste', we check if an image was loaded before/after to decide on undo tagging.

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

After handling the main edit actions, <SwmToken path="Modules/Processor.bas" pos="1548:1:1" line-data="        Process_EditMenu = True">`Process_EditMenu`</SwmToken> finishes up by handling special clipboard actions (with dialogs if needed), selection filters, and syncing <SwmToken path="Modules/Processor.bas" pos="1608:6:8" line-data="        &#39;Synchronize any non-destructive settings to the currently active layer">`non-destructive`</SwmToken> layer settings if undo/redo was used.

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

## Image Menu Dispatch and Post-Edit Operations

<SwmSnippet path="/Modules/Processor.bas" line="243">

---

Back in <SwmToken path="Modules/Processor.bas" pos="89:4:4" line-data="Public Sub Process(ByVal processID As String, Optional raiseDialog As Boolean = False, Optional processParameters As String = vbNullString, Optional createUndo As PD_UndoType = UNDO_Nothing, Optional relevantTool As Long = -1, Optional recordAction As Boolean = True)">`Process`</SwmToken>, if the <SwmToken path="Modules/Processor.bas" pos="244:17:17" line-data="    If (Not processFound) Then processFound = Process_ImageMenu(processID, raiseDialog, processParameters, createUndo, relevantTool, recordAction, returnDetails)">`processID`</SwmToken> wasn't handled by file or edit menus, we check if it's an image menu action and dispatch to <SwmToken path="Modules/Processor.bas" pos="244:15:15" line-data="    If (Not processFound) Then processFound = Process_ImageMenu(processID, raiseDialog, processParameters, createUndo, relevantTool, recordAction, returnDetails)">`Process_ImageMenu`</SwmToken> next.

```visual basic
    'Image menu operations have been successfully migrated to XML strings.  (None of their functions raise special return conditions, FYI.)
    If (Not processFound) Then processFound = Process_ImageMenu(processID, raiseDialog, processParameters, createUndo, relevantTool, recordAction, returnDetails)
    
```

---

</SwmSnippet>

## Image Menu Command Handling and Dialog Routing

```mermaid
%%{init: {"flowchart": {"defaultRenderer": "elk"}} }%%
flowchart TD
    node1["User requests an image menu action"] --> node2{"Is the requested action (processID) supported?"}
    click node1 openCode "Modules/Processor.bas:2204:2335"
    node2 -->|"Supported"| node3{"Should a dialog be shown? (raiseDialog)"}
    click node2 openCode "Modules/Processor.bas:2209:2332"
    node3 -->|"Yes"| node4["Show relevant dialog for image operation"]
    click node3 openCode "Modules/Processor.bas:2216:2332"
    node3 -->|"No"| node5["Perform image operation directly"]
    click node4 openCode "Modules/Processor.bas:2216:2332"
    click node5 openCode "Modules/Processor.bas:2216:2332"
    node2 -->|"Legacy/Removed"| node6["Do nothing, return success"]
    click node6 openCode "Modules/Processor.bas:2317:2332"

classDef HeadingStyle fill:#777777,stroke:#333,stroke-width:2px;

%% Swimm:
%% %%{init: {"flowchart": {"defaultRenderer": "elk"}} }%%
%% flowchart TD
%%     node1["User requests an image menu action"] --> node2{"Is the requested action (<SwmToken path="Modules/Processor.bas" pos="89:8:8" line-data="Public Sub Process(ByVal processID As String, Optional raiseDialog As Boolean = False, Optional processParameters As String = vbNullString, Optional createUndo As PD_UndoType = UNDO_Nothing, Optional relevantTool As Long = -1, Optional recordAction As Boolean = True)">`processID`</SwmToken>) supported?"}
%%     click node1 openCode "<SwmPath>[Modules/Processor.bas](Modules/Processor.bas)</SwmPath>:2204:2335"
%%     node2 -->|"Supported"| node3{"Should a dialog be shown? (<SwmToken path="Modules/Processor.bas" pos="89:17:17" line-data="Public Sub Process(ByVal processID As String, Optional raiseDialog As Boolean = False, Optional processParameters As String = vbNullString, Optional createUndo As PD_UndoType = UNDO_Nothing, Optional relevantTool As Long = -1, Optional recordAction As Boolean = True)">`raiseDialog`</SwmToken>)"}
%%     click node2 openCode "<SwmPath>[Modules/Processor.bas](Modules/Processor.bas)</SwmPath>:2209:2332"
%%     node3 -->|"Yes"| node4["Show relevant dialog for image operation"]
%%     click node3 openCode "<SwmPath>[Modules/Processor.bas](Modules/Processor.bas)</SwmPath>:2216:2332"
%%     node3 -->|"No"| node5["Perform image operation directly"]
%%     click node4 openCode "<SwmPath>[Modules/Processor.bas](Modules/Processor.bas)</SwmPath>:2216:2332"
%%     click node5 openCode "<SwmPath>[Modules/Processor.bas](Modules/Processor.bas)</SwmPath>:2216:2332"
%%     node2 -->|"Legacy/Removed"| node6["Do nothing, return success"]
%%     click node6 openCode "<SwmPath>[Modules/Processor.bas](Modules/Processor.bas)</SwmPath>:2317:2332"
%% 
%% classDef HeadingStyle fill:#777777,stroke:#333,stroke-width:2px;
```

<SwmSnippet path="/Modules/Processor.bas" line="2204">

---

In <SwmToken path="Modules/Processor.bas" pos="2204:4:4" line-data="Private Function Process_ImageMenu(ByVal processID As String, Optional raiseDialog As Boolean = False, Optional processParameters As String = vbNullString, Optional createUndo As PD_UndoType = UNDO_Nothing, Optional relevantTool As Long = -1, Optional recordAction As Boolean = True, Optional ByRef returnDetails As String = vbNullString) As Boolean">`Process_ImageMenu`</SwmToken>, we match the <SwmToken path="Modules/Processor.bas" pos="2204:8:8" line-data="Private Function Process_ImageMenu(ByVal processID As String, Optional raiseDialog As Boolean = False, Optional processParameters As String = vbNullString, Optional createUndo As PD_UndoType = UNDO_Nothing, Optional relevantTool As Long = -1, Optional recordAction As Boolean = True, Optional ByRef returnDetails As String = vbNullString) As Boolean">`processID`</SwmToken> against known image commands. For dialog-based actions (like 'Resize image'), if <SwmToken path="Modules/Processor.bas" pos="2204:17:17" line-data="Private Function Process_ImageMenu(ByVal processID As String, Optional raiseDialog As Boolean = False, Optional processParameters As String = vbNullString, Optional createUndo As PD_UndoType = UNDO_Nothing, Optional relevantTool As Long = -1, Optional recordAction As Boolean = True, Optional ByRef returnDetails As String = vbNullString) As Boolean">`raiseDialog`</SwmToken> is set, we call dialog functions (like <SwmToken path="Modules/Processor.bas" pos="2216:7:7" line-data="        If raiseDialog Then ShowResizeDialog pdat_Image Else FormResize.ResizeImage processParameters">`ShowResizeDialog`</SwmToken> in <SwmPath>[Modules/DialogManager.bas](Modules/DialogManager.bas)</SwmPath>); otherwise, we run the action directly. This pattern repeats for each image menu command, and legacy commands are handled as no-ops for macro compatibility.

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

<SwmToken path="Modules/DialogManager.bas" pos="324:4:4" line-data="Public Sub ShowResizeDialog(ByVal ResizeTarget As PD_ActionTarget)">`ShowResizeDialog`</SwmToken> sets the target for resizing on the <SwmToken path="Modules/DialogManager.bas" pos="325:1:1" line-data="    FormResize.ResizeTarget = ResizeTarget">`FormResize`</SwmToken> dialog, then shows it modally using <SwmToken path="Modules/DialogManager.bas" pos="326:1:1" line-data="    ShowPDDialog vbModal, FormResize">`ShowPDDialog`</SwmToken>. This lets the user interactively set resize options for the right target.

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

Back in <SwmToken path="Modules/Processor.bas" pos="2221:1:1" line-data="        Process_ImageMenu = True">`Process_ImageMenu`</SwmToken>, if the command is <SwmToken path="Modules/Interface.bas" pos="847:4:6" line-data="            &#39;The content-aware fill option in the edit menu also requires an active selection.">`content-aware`</SwmToken> resize and <SwmToken path="Modules/Processor.bas" pos="2220:3:3" line-data="        If raiseDialog Then ShowContentAwareResizeDialog pdat_Image Else FormResizeContentAware.SmartResizeImage processParameters">`raiseDialog`</SwmToken> is set, we call <SwmToken path="Modules/Processor.bas" pos="2220:7:7" line-data="        If raiseDialog Then ShowContentAwareResizeDialog pdat_Image Else FormResizeContentAware.SmartResizeImage processParameters">`ShowContentAwareResizeDialog`</SwmToken> to open the right dialog for that operation.

```visual basic
        If raiseDialog Then ShowContentAwareResizeDialog pdat_Image Else FormResizeContentAware.SmartResizeImage processParameters
        Process_ImageMenu = True
        
    ElseIf Strings.StringsEqual(processID, "Canvas size", True) Then
```

---

</SwmSnippet>

<SwmSnippet path="/Modules/DialogManager.bas" line="330">

---

<SwmToken path="Modules/DialogManager.bas" pos="330:4:4" line-data="Public Sub ShowContentAwareResizeDialog(ByVal ResizeTarget As PD_ActionTarget)">`ShowContentAwareResizeDialog`</SwmToken> sets the target for the <SwmToken path="Modules/Interface.bas" pos="847:4:6" line-data="            &#39;The content-aware fill option in the edit menu also requires an active selection.">`content-aware`</SwmToken> resize dialog, then shows it modally. This keeps the dialog focused on the right image/object.

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

Back in <SwmToken path="Modules/Processor.bas" pos="2225:1:1" line-data="        Process_ImageMenu = True">`Process_ImageMenu`</SwmToken>, for canvas size changes, if <SwmToken path="Modules/Processor.bas" pos="2224:3:3" line-data="        If raiseDialog Then ShowPDDialog vbModal, FormCanvasSize Else FormCanvasSize.ResizeCanvas processParameters">`raiseDialog`</SwmToken> is set, we show the <SwmToken path="Modules/Processor.bas" pos="2224:12:12" line-data="        If raiseDialog Then ShowPDDialog vbModal, FormCanvasSize Else FormCanvasSize.ResizeCanvas processParameters">`FormCanvasSize`</SwmToken> dialog modally (<SwmToken path="Modules/Processor.bas" pos="2224:7:7" line-data="        If raiseDialog Then ShowPDDialog vbModal, FormCanvasSize Else FormCanvasSize.ResizeCanvas processParameters">`ShowPDDialog`</SwmToken>); otherwise, we run the resize directly.

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

Back in <SwmToken path="Modules/Processor.bas" pos="2239:1:1" line-data="        Process_ImageMenu = True">`Process_ImageMenu`</SwmToken>, for straighten image, if <SwmToken path="Modules/Processor.bas" pos="2235:27:27" line-data="    &#39;Crop operations.  Note that the main form submits &quot;Crop&quot; requests with raiseDialog set to TRUE.  This tells us to ask the">`raiseDialog`</SwmToken> is set, we call <SwmToken path="Modules/Processor.bas" pos="2246:7:7" line-data="        If raiseDialog Then ShowStraightenDialog pdat_Image Else FormStraighten.StraightenImage processParameters">`ShowStraightenDialog`</SwmToken> to let the user pick the angle. Otherwise, we run the straighten directly.

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

<SwmToken path="Modules/DialogManager.bas" pos="342:4:4" line-data="Public Sub ShowStraightenDialog(ByVal StraightenTarget As PD_ActionTarget)">`ShowStraightenDialog`</SwmToken> sets the target for the straighten operation on the dialog, then shows it modally so the user can interactively set the angle.

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

Back in <SwmToken path="Modules/Processor.bas" pos="2263:1:1" line-data="        Process_ImageMenu = True">`Process_ImageMenu`</SwmToken>, for arbitrary rotation, if <SwmToken path="Modules/Processor.bas" pos="2262:3:3" line-data="        If raiseDialog Then ShowRotateDialog pdat_Image Else FormRotate.RotateArbitrary processParameters">`raiseDialog`</SwmToken> is set, we call <SwmToken path="Modules/Processor.bas" pos="2262:7:7" line-data="        If raiseDialog Then ShowRotateDialog pdat_Image Else FormRotate.RotateArbitrary processParameters">`ShowRotateDialog`</SwmToken> to let the user pick the angle. Otherwise, we run the rotation directly.

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

<SwmToken path="Modules/DialogManager.bas" pos="336:4:4" line-data="Public Sub ShowRotateDialog(ByVal RotateTarget As PD_ActionTarget)">`ShowRotateDialog`</SwmToken> sets the target for the rotation operation on the dialog, then shows it modally so the user can interactively set the angle.

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

Back in <SwmToken path="Modules/Processor.bas" pos="2276:1:1" line-data="        Process_ImageMenu = True">`Process_ImageMenu`</SwmToken>, for flatten image, if <SwmToken path="Modules/Processor.bas" pos="2282:3:3" line-data="        If raiseDialog Then">`raiseDialog`</SwmToken> is set and the image has transparency, we show the flatten dialog. Otherwise, we flatten directly. This avoids unnecessary dialogs when not needed.

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

After all the dialog and direct action handling, <SwmToken path="Modules/Processor.bas" pos="2307:1:1" line-data="        Process_ImageMenu = True">`Process_ImageMenu`</SwmToken> finishes up by handling legacy and removed commands as no-ops. This keeps old macros from breaking if they reference commands that no longer exist.

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

## Layer Menu Dispatch and Post-Image Operations

<SwmSnippet path="/Modules/Processor.bas" line="246">

---

Back in <SwmToken path="Modules/Processor.bas" pos="89:4:4" line-data="Public Sub Process(ByVal processID As String, Optional raiseDialog As Boolean = False, Optional processParameters As String = vbNullString, Optional createUndo As PD_UndoType = UNDO_Nothing, Optional relevantTool As Long = -1, Optional recordAction As Boolean = True)">`Process`</SwmToken>, if the <SwmToken path="Modules/Processor.bas" pos="247:17:17" line-data="    If (Not processFound) Then processFound = Process_LayerMenu(processID, raiseDialog, processParameters, createUndo, relevantTool, recordAction, returnDetails)">`processID`</SwmToken> wasn't handled by file, edit, or image menus, we check if it's a layer menu action and dispatch to <SwmToken path="Modules/Processor.bas" pos="247:15:15" line-data="    If (Not processFound) Then processFound = Process_LayerMenu(processID, raiseDialog, processParameters, createUndo, relevantTool, recordAction, returnDetails)">`Process_LayerMenu`</SwmToken> next.

```visual basic
    'Layer menu operations have been successfully migrated to XML strings.  (None of their functions raise special return conditions, FYI.)
    If (Not processFound) Then processFound = Process_LayerMenu(processID, raiseDialog, processParameters, createUndo, relevantTool, recordAction, returnDetails)
    
```

---

</SwmSnippet>

## Layer Operations Dispatch

```mermaid
%%{init: {"flowchart": {"defaultRenderer": "elk"}} }%%
flowchart TD
    node1["User requests a layer operation (processID)"] --> node2{"What type of operation?"}
    click node1 openCode "Modules/Processor.bas:2340:2347"
    node2 -->|"Add/duplicate layer"| node3["Add or duplicate a layer"]
    click node2 openCode "Modules/Processor.bas:2349:2390"
    node2 -->|"Delete layer"| node4["Remove a layer"]
    click node4 openCode "Modules/Processor.bas:2405:2412"
    node2 -->|"Merge/move/reorder layer"| node5["Change layer order or position"]
    click node5 openCode "Modules/Processor.bas:2427:2474"
    node2 -->|"Transform/resize/crop"| node6{"Does this require user input?"}
    click node6 openCode "Modules/Processor.bas:2499:2552"
    node6 -->|"Yes"| node7["Show dialog for user input"]
    click node7 openCode "Modules/Processor.bas:2543:2544"
    node7 --> node8["Perform transformation"]
    click node8 openCode "Modules/Processor.bas:2543:2544"
    node6 -->|"No"| node8
    node2 -->|"Visibility"| node9["Show or hide layers"]
    click node9 openCode "Modules/Processor.bas:2475:2497"
    node2 -->|"Alpha/transparency"| node10["Adjust transparency"]
    click node10 openCode "Modules/Processor.bas:2563:2579"
    node2 -->|"Split/rasterize/rearrange"| node11["Convert or rearrange layers"]
    click node11 openCode "Modules/Processor.bas:2581:2626"
    node3 --> node12["Operation complete"]
    node4 --> node12
    node5 --> node12
    node8 --> node12
    node9 --> node12
    node10 --> node12
    node11 --> node12
    click node12 openCode "Modules/Processor.bas:2340:2628"
classDef HeadingStyle fill:#777777,stroke:#333,stroke-width:2px;

%% Swimm:
%% %%{init: {"flowchart": {"defaultRenderer": "elk"}} }%%
%% flowchart TD
%%     node1["User requests a layer operation (<SwmToken path="Modules/Processor.bas" pos="89:8:8" line-data="Public Sub Process(ByVal processID As String, Optional raiseDialog As Boolean = False, Optional processParameters As String = vbNullString, Optional createUndo As PD_UndoType = UNDO_Nothing, Optional relevantTool As Long = -1, Optional recordAction As Boolean = True)">`processID`</SwmToken>)"] --> node2{"What type of operation?"}
%%     click node1 openCode "<SwmPath>[Modules/Processor.bas](Modules/Processor.bas)</SwmPath>:2340:2347"
%%     node2 -->|"Add/duplicate layer"| node3["Add or duplicate a layer"]
%%     click node2 openCode "<SwmPath>[Modules/Processor.bas](Modules/Processor.bas)</SwmPath>:2349:2390"
%%     node2 -->|"Delete layer"| node4["Remove a layer"]
%%     click node4 openCode "<SwmPath>[Modules/Processor.bas](Modules/Processor.bas)</SwmPath>:2405:2412"
%%     node2 -->|"Merge/move/reorder layer"| node5["Change layer order or position"]
%%     click node5 openCode "<SwmPath>[Modules/Processor.bas](Modules/Processor.bas)</SwmPath>:2427:2474"
%%     node2 -->|"Transform/resize/crop"| node6{"Does this require user input?"}
%%     click node6 openCode "<SwmPath>[Modules/Processor.bas](Modules/Processor.bas)</SwmPath>:2499:2552"
%%     node6 -->|"Yes"| node7["Show dialog for user input"]
%%     click node7 openCode "<SwmPath>[Modules/Processor.bas](Modules/Processor.bas)</SwmPath>:2543:2544"
%%     node7 --> node8["Perform transformation"]
%%     click node8 openCode "<SwmPath>[Modules/Processor.bas](Modules/Processor.bas)</SwmPath>:2543:2544"
%%     node6 -->|"No"| node8
%%     node2 -->|"Visibility"| node9["Show or hide layers"]
%%     click node9 openCode "<SwmPath>[Modules/Processor.bas](Modules/Processor.bas)</SwmPath>:2475:2497"
%%     node2 -->|"Alpha/transparency"| node10["Adjust transparency"]
%%     click node10 openCode "<SwmPath>[Modules/Processor.bas](Modules/Processor.bas)</SwmPath>:2563:2579"
%%     node2 -->|"Split/rasterize/rearrange"| node11["Convert or rearrange layers"]
%%     click node11 openCode "<SwmPath>[Modules/Processor.bas](Modules/Processor.bas)</SwmPath>:2581:2626"
%%     node3 --> node12["Operation complete"]
%%     node4 --> node12
%%     node5 --> node12
%%     node8 --> node12
%%     node9 --> node12
%%     node10 --> node12
%%     node11 --> node12
%%     click node12 openCode "<SwmPath>[Modules/Processor.bas](Modules/Processor.bas)</SwmPath>:2340:2628"
%% classDef HeadingStyle fill:#777777,stroke:#333,stroke-width:2px;
```

<SwmSnippet path="/Modules/Processor.bas" line="2340">

---

In <SwmToken path="Modules/Processor.bas" pos="2340:4:4" line-data="Private Function Process_LayerMenu(ByVal processID As String, Optional ByVal raiseDialog As Boolean = False, Optional ByRef processParameters As String = vbNullString, Optional ByRef createUndo As PD_UndoType = UNDO_Nothing, Optional ByVal relevantTool As Long = -1, Optional ByRef recordAction As Boolean = True, Optional ByRef returnDetails As String = vbNullString) As Boolean">`Process_LayerMenu`</SwmToken>, we start by parsing any parameters with <SwmToken path="Modules/Processor.bas" pos="2344:7:7" line-data="    Dim cParams As pdSerialize">`pdSerialize`</SwmToken> so we can handle both simple and complex commands. The function then runs through a big If-ElseIf block, matching the <SwmToken path="Modules/Processor.bas" pos="2340:8:8" line-data="Private Function Process_LayerMenu(ByVal processID As String, Optional ByVal raiseDialog As Boolean = False, Optional ByRef processParameters As String = vbNullString, Optional ByRef createUndo As PD_UndoType = UNDO_Nothing, Optional ByVal relevantTool As Long = -1, Optional ByRef recordAction As Boolean = True, Optional ByRef returnDetails As String = vbNullString) As Boolean">`processID`</SwmToken> to the right layer operation—like adding, duplicating, merging, or toggling visibility. For commands that need user input, we call <SwmToken path="Modules/Processor.bas" pos="2354:7:7" line-data="        If raiseDialog Then ShowPDDialog vbModal, FormNewLayer Else Layers.AddNewLayer_XML processParameters">`ShowPDDialog`</SwmToken> (from <SwmPath>[Modules/Interface.bas](Modules/Interface.bas)</SwmPath>) to pop up the right dialog. This setup keeps all layer actions in one place and makes it easy to add new ones or tweak parameter handling later.

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

After returning from <SwmPath>[Modules/Interface.bas](Modules/Interface.bas)</SwmPath> (for dialogs), <SwmToken path="Modules/Processor.bas" pos="2382:1:1" line-data="        Process_LayerMenu = True">`Process_LayerMenu`</SwmToken> checks if we're handling a text or typography layer during macro playback or batch mode. If so, it creates the new layer and restores its state from XML, so macros and batch jobs can recreate text layers exactly as recorded. For other layer commands that need dialogs (like 'New layer from file'), we call DialogManager next to show the right UI.

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

After coming back from <SwmPath>[Modules/DialogManager.bas](Modules/DialogManager.bas)</SwmPath> (for things like 'Straighten layer'), <SwmToken path="Modules/Processor.bas" pos="2544:1:1" line-data="        Process_LayerMenu = True">`Process_LayerMenu`</SwmToken> checks the <SwmToken path="Modules/Processor.bas" pos="2543:3:3" line-data="        If raiseDialog Then ShowRotateDialog pdat_SingleLayer Else FormRotate.RotateArbitrary processParameters">`raiseDialog`</SwmToken> flag—if it's set, we show the dialog again for the next operation (like arbitrary rotation); otherwise, we run the action directly. This keeps the UI and automation paths in sync.

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

After <SwmPath>[Modules/DialogManager.bas](Modules/DialogManager.bas)</SwmPath> (for arbitrary rotation), <SwmToken path="Modules/Processor.bas" pos="2557:1:1" line-data="        Process_LayerMenu = True">`Process_LayerMenu`</SwmToken> checks if the next operation (like resize or <SwmToken path="Modules/Interface.bas" pos="847:4:6" line-data="            &#39;The content-aware fill option in the edit menu also requires an active selection.">`content-aware`</SwmToken> resize) needs a dialog. If <SwmToken path="Modules/Processor.bas" pos="2556:3:3" line-data="        If raiseDialog Then ShowResizeDialog pdat_SingleLayer Else FormResize.ResizeImage processParameters">`raiseDialog`</SwmToken> is True, we call DialogManager again to show the right UI; otherwise, we just run the resize directly. This keeps the workflow consistent for both interactive and automated use.

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

After returning from <SwmPath>[Modules/DialogManager.bas](Modules/DialogManager.bas)</SwmPath> (for resize), <SwmToken path="Modules/Processor.bas" pos="2561:1:1" line-data="        Process_LayerMenu = True">`Process_LayerMenu`</SwmToken> checks if the next operation (like <SwmToken path="Modules/Interface.bas" pos="847:4:6" line-data="            &#39;The content-aware fill option in the edit menu also requires an active selection.">`content-aware`</SwmToken> resize) needs a dialog. If so, DialogManager is called again to show the right UI for that action; otherwise, the operation runs directly. This keeps each layer action modular and user-friendly.

```visual basic
        If raiseDialog Then ShowContentAwareResizeDialog pdat_SingleLayer Else FormResizeContentAware.SmartResizeImage processParameters
        Process_LayerMenu = True
        
```

---

</SwmSnippet>

<SwmSnippet path="/Modules/Processor.bas" line="2563">

---

After <SwmPath>[Modules/DialogManager.bas](Modules/DialogManager.bas)</SwmPath> (for <SwmToken path="Modules/Interface.bas" pos="847:4:6" line-data="            &#39;The content-aware fill option in the edit menu also requires an active selection.">`content-aware`</SwmToken> resize), <SwmToken path="Modules/Processor.bas" pos="2566:1:1" line-data="        Process_LayerMenu = True">`Process_LayerMenu`</SwmToken> moves on to alpha/transparency commands. If <SwmToken path="Modules/Processor.bas" pos="2565:3:3" line-data="        If raiseDialog Then ShowPDDialog vbModal, FormTransparency_FromColor Else FormTransparency_FromColor.ColorToAlpha processParameters">`raiseDialog`</SwmToken> is set, we call <SwmPath>[Modules/Interface.bas](Modules/Interface.bas)</SwmPath> to show the right dialog (like <SwmToken path="Modules/Processor.bas" pos="2565:12:12" line-data="        If raiseDialog Then ShowPDDialog vbModal, FormTransparency_FromColor Else FormTransparency_FromColor.ColorToAlpha processParameters">`FormTransparency_FromColor`</SwmToken>); otherwise, we run the operation directly. This keeps transparency actions consistent with the rest of the UI.

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

After <SwmPath>[Modules/Interface.bas](Modules/Interface.bas)</SwmPath> (for dialogs like splitting layers/images), <SwmToken path="Modules/Processor.bas" pos="2596:1:1" line-data="        Process_LayerMenu = True">`Process_LayerMenu`</SwmToken> wraps up by handling the rest of the layer commands—like rasterizing, <SwmToken path="Modules/Processor.bas" pos="2603:16:18" line-data="    ElseIf Strings.StringsEqual(processID, &quot;Resize layer (on-canvas)&quot;, True) Then">`on-canvas`</SwmToken> moves, and dummy entries for undo/redo. The dispatcher pattern here means every layer action is routed through this one function, so adding or changing commands is straightforward.

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

## Selection Operations Dispatch

```mermaid
%%{init: {"flowchart": {"defaultRenderer": "elk"}} }%%
flowchart TD
    node1["Start processing operation"] --> node2{"Has operation already been handled? (processFound)"}
    click node1 openCode "Modules/Processor.bas:249:249"
    click node2 openCode "Modules/Processor.bas:249:249"
    node2 -->|"No"| node3["Attempt selection menu operation (processID)"]
    click node3 openCode "Modules/Processor.bas:250:250"
    node3 --> node4{"Was operation handled? (processFound)"}
    click node4 openCode "Modules/Processor.bas:250:250"
    node4 -->|"No"| node5["Attempt adjustment menu operation (processID)"]
    click node5 openCode "Modules/Processor.bas:253:253"
    node4 -->|"Yes"| node6["Done"]
    node5 --> node6["Done"]
    node2 -->|"Yes"| node6["Done"]
    click node6 openCode "Modules/Processor.bas:253:253"

classDef HeadingStyle fill:#777777,stroke:#333,stroke-width:2px;

%% Swimm:
%% %%{init: {"flowchart": {"defaultRenderer": "elk"}} }%%
%% flowchart TD
%%     node1["Start processing operation"] --> node2{"Has operation already been handled? (<SwmToken path="Modules/Processor.bas" pos="216:3:3" line-data="    Dim processFound As Boolean, returnDetails As String">`processFound`</SwmToken>)"}
%%     click node1 openCode "<SwmPath>[Modules/Processor.bas](Modules/Processor.bas)</SwmPath>:249:249"
%%     click node2 openCode "<SwmPath>[Modules/Processor.bas](Modules/Processor.bas)</SwmPath>:249:249"
%%     node2 -->|"No"| node3["Attempt selection menu operation (<SwmToken path="Modules/Processor.bas" pos="89:8:8" line-data="Public Sub Process(ByVal processID As String, Optional raiseDialog As Boolean = False, Optional processParameters As String = vbNullString, Optional createUndo As PD_UndoType = UNDO_Nothing, Optional relevantTool As Long = -1, Optional recordAction As Boolean = True)">`processID`</SwmToken>)"]
%%     click node3 openCode "<SwmPath>[Modules/Processor.bas](Modules/Processor.bas)</SwmPath>:250:250"
%%     node3 --> node4{"Was operation handled? (<SwmToken path="Modules/Processor.bas" pos="216:3:3" line-data="    Dim processFound As Boolean, returnDetails As String">`processFound`</SwmToken>)"}
%%     click node4 openCode "<SwmPath>[Modules/Processor.bas](Modules/Processor.bas)</SwmPath>:250:250"
%%     node4 -->|"No"| node5["Attempt adjustment menu operation (<SwmToken path="Modules/Processor.bas" pos="89:8:8" line-data="Public Sub Process(ByVal processID As String, Optional raiseDialog As Boolean = False, Optional processParameters As String = vbNullString, Optional createUndo As PD_UndoType = UNDO_Nothing, Optional relevantTool As Long = -1, Optional recordAction As Boolean = True)">`processID`</SwmToken>)"]
%%     click node5 openCode "<SwmPath>[Modules/Processor.bas](Modules/Processor.bas)</SwmPath>:253:253"
%%     node4 -->|"Yes"| node6["Done"]
%%     node5 --> node6["Done"]
%%     node2 -->|"Yes"| node6["Done"]
%%     click node6 openCode "<SwmPath>[Modules/Processor.bas](Modules/Processor.bas)</SwmPath>:253:253"
%% 
%% classDef HeadingStyle fill:#777777,stroke:#333,stroke-width:2px;
```

<SwmSnippet path="/Modules/Processor.bas" line="249">

---

Back in Process, after returning from <SwmToken path="Modules/Processor.bas" pos="247:15:15" line-data="    If (Not processFound) Then processFound = Process_LayerMenu(processID, raiseDialog, processParameters, createUndo, relevantTool, recordAction, returnDetails)">`Process_LayerMenu`</SwmToken>, we check if the <SwmToken path="Modules/Processor.bas" pos="250:17:17" line-data="    If (Not processFound) Then processFound = Process_SelectMenu(processID, raiseDialog, processParameters, createUndo, relevantTool, recordAction, returnDetails)">`processID`</SwmToken> matches any selection menu actions. If not already handled, we call <SwmToken path="Modules/Processor.bas" pos="250:15:15" line-data="    If (Not processFound) Then processFound = Process_SelectMenu(processID, raiseDialog, processParameters, createUndo, relevantTool, recordAction, returnDetails)">`Process_SelectMenu`</SwmToken> next to see if it's a selection-related command. This keeps the flow moving through each menu type in order.

```visual basic
    'Select menu operations have been successfully migrated to XML strings.  (None of their functions raise special return conditions, FYI.)
    If (Not processFound) Then processFound = Process_SelectMenu(processID, raiseDialog, processParameters, createUndo, relevantTool, recordAction, returnDetails)
    
```

---

</SwmSnippet>

<SwmSnippet path="/Modules/Processor.bas" line="2633">

---

<SwmToken path="Modules/Processor.bas" pos="2633:4:4" line-data="Private Function Process_SelectMenu(ByVal processID As String, Optional raiseDialog As Boolean = False, Optional processParameters As String = vbNullString, Optional createUndo As PD_UndoType = UNDO_Nothing, Optional relevantTool As Long = -1, Optional recordAction As Boolean = True, Optional ByRef returnDetails As String = vbNullString) As Boolean">`Process_SelectMenu`</SwmToken> parses parameters with <SwmToken path="Modules/Processor.bas" pos="2637:7:7" line-data="    Dim cParams As pdSerialize">`pdSerialize`</SwmToken> and then checks the <SwmToken path="Modules/Processor.bas" pos="2633:8:8" line-data="Private Function Process_SelectMenu(ByVal processID As String, Optional raiseDialog As Boolean = False, Optional processParameters As String = vbNullString, Optional createUndo As PD_UndoType = UNDO_Nothing, Optional relevantTool As Long = -1, Optional recordAction As Boolean = True, Optional ByRef returnDetails As String = vbNullString) As Boolean">`processID`</SwmToken> against a bunch of selection actions—like create, remove, invert, grow, shrink, feather, and more. Some actions use <SwmToken path="Modules/Processor.bas" pos="2633:17:17" line-data="Private Function Process_SelectMenu(ByVal processID As String, Optional raiseDialog As Boolean = False, Optional processParameters As String = vbNullString, Optional createUndo As PD_UndoType = UNDO_Nothing, Optional relevantTool As Long = -1, Optional recordAction As Boolean = True, Optional ByRef returnDetails As String = vbNullString) As Boolean">`raiseDialog`</SwmToken> to show a UI, others just use the parameters. Dummy entries for move/resize selection exist just to keep undo/redo working right.

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

Back in Process, after returning from <SwmToken path="Modules/Processor.bas" pos="250:15:15" line-data="    If (Not processFound) Then processFound = Process_SelectMenu(processID, raiseDialog, processParameters, createUndo, relevantTool, recordAction, returnDetails)">`Process_SelectMenu`</SwmToken>, we check if the <SwmToken path="Modules/Processor.bas" pos="253:17:17" line-data="    If (Not processFound) Then processFound = Process_AdjustmentsMenu(processID, raiseDialog, processParameters, createUndo, relevantTool, recordAction, returnDetails)">`processID`</SwmToken> matches any adjustment menu actions. If not already handled, we call <SwmToken path="Modules/Processor.bas" pos="253:15:15" line-data="    If (Not processFound) Then processFound = Process_AdjustmentsMenu(processID, raiseDialog, processParameters, createUndo, relevantTool, recordAction, returnDetails)">`Process_AdjustmentsMenu`</SwmToken> next to see if it's an adjustment-related command. This keeps the flow modular and lets each handler focus on its own set of actions.

```visual basic
    'Adjustment menu operations have been successfully migrated to XML strings.  (None of their functions raise special return conditions, FYI.)
    If (Not processFound) Then processFound = Process_AdjustmentsMenu(processID, raiseDialog, processParameters, createUndo, relevantTool, recordAction, returnDetails)
    
```

---

</SwmSnippet>

## Image Adjustments Dispatch

```mermaid
%%{init: {"flowchart": {"defaultRenderer": "elk"}} }%%
flowchart TD
    node1["User requests image adjustment"] --> node2{"Which adjustment is selected?"}
    click node1 openCode "Modules/Processor.bas:2030:2127"
    node2 -->|"Film negative"| node3["Film Negative Filter and Progress Updates"]
    click node2 openCode "Modules/Processor.bas:2128:2132"
    
    node2 -->|"Invert hue"| node4["Dispatching and Running Hue Inversion"]
    
    node2 -->|"Invert RGB"| node5["Dispatching and Running RGB Inversion"]
    
    node2 -->|"Shift colors"| node6{"Shift direction?"}
    click node6 openCode "Modules/Processor.bas:2168:2174"
    node6 -->|"Left"| node7["Dispatching and Running Color Channel Shift"]
    
    node6 -->|"Right"| node8["Dispatching and Running Color Channel Shift"]
    
    node2 -->|"Maximum channel"| node9["Max/Min Channel Filtering and Calculation"]
    
    node2 -->|"Minimum channel"| node10["Max/Min Channel Filtering and Calculation"]
    
    node2 -->|"Other adjustment"| node11{"Show dialog or apply automatically?"}
    click node11 openCode "Modules/Processor.bas:2030:2184"
    node11 -->|"Show dialog"| node12["Show adjustment dialog"]
    click node12 openCode "Modules/Processor.bas:2042:2042"
    node11 -->|"Apply automatically"| node13["Apply adjustment automatically"]
    click node13 openCode "Modules/Processor.bas:2042:2042"
classDef HeadingStyle fill:#777777,stroke:#333,stroke-width:2px;
click node3 goToHeading "Film Negative Filter and Progress Updates"
node3:::HeadingStyle
click node4 goToHeading "Dispatching and Running Hue Inversion"
node4:::HeadingStyle
click node5 goToHeading "Dispatching and Running RGB Inversion"
node5:::HeadingStyle
click node7 goToHeading "Dispatching and Running Color Channel Shift"
node7:::HeadingStyle
click node8 goToHeading "Dispatching and Running Color Channel Shift"
node8:::HeadingStyle
click node9 goToHeading "Max/Min Channel Filtering and Calculation"
node9:::HeadingStyle
click node10 goToHeading "Max/Min Channel Filtering and Calculation"
node10:::HeadingStyle

%% Swimm:
%% %%{init: {"flowchart": {"defaultRenderer": "elk"}} }%%
%% flowchart TD
%%     node1["User requests image adjustment"] --> node2{"Which adjustment is selected?"}
%%     click node1 openCode "<SwmPath>[Modules/Processor.bas](Modules/Processor.bas)</SwmPath>:2030:2127"
%%     node2 -->|"Film negative"| node3["Film Negative Filter and Progress Updates"]
%%     click node2 openCode "<SwmPath>[Modules/Processor.bas](Modules/Processor.bas)</SwmPath>:2128:2132"
%%     
%%     node2 -->|"Invert hue"| node4["Dispatching and Running Hue Inversion"]
%%     
%%     node2 -->|"Invert RGB"| node5["Dispatching and Running RGB Inversion"]
%%     
%%     node2 -->|"Shift colors"| node6{"Shift direction?"}
%%     click node6 openCode "<SwmPath>[Modules/Processor.bas](Modules/Processor.bas)</SwmPath>:2168:2174"
%%     node6 -->|"Left"| node7["Dispatching and Running Color Channel Shift"]
%%     
%%     node6 -->|"Right"| node8["Dispatching and Running Color Channel Shift"]
%%     
%%     node2 -->|"Maximum channel"| node9["Max/Min Channel Filtering and Calculation"]
%%     
%%     node2 -->|"Minimum channel"| node10["Max/Min Channel Filtering and Calculation"]
%%     
%%     node2 -->|"Other adjustment"| node11{"Show dialog or apply automatically?"}
%%     click node11 openCode "<SwmPath>[Modules/Processor.bas](Modules/Processor.bas)</SwmPath>:2030:2184"
%%     node11 -->|"Show dialog"| node12["Show adjustment dialog"]
%%     click node12 openCode "<SwmPath>[Modules/Processor.bas](Modules/Processor.bas)</SwmPath>:2042:2042"
%%     node11 -->|"Apply automatically"| node13["Apply adjustment automatically"]
%%     click node13 openCode "<SwmPath>[Modules/Processor.bas](Modules/Processor.bas)</SwmPath>:2042:2042"
%% classDef HeadingStyle fill:#777777,stroke:#333,stroke-width:2px;
%% click node3 goToHeading "Film Negative Filter and Progress Updates"
%% node3:::HeadingStyle
%% click node4 goToHeading "Dispatching and Running Hue Inversion"
%% node4:::HeadingStyle
%% click node5 goToHeading "Dispatching and Running RGB Inversion"
%% node5:::HeadingStyle
%% click node7 goToHeading "Dispatching and Running Color Channel Shift"
%% node7:::HeadingStyle
%% click node8 goToHeading "Dispatching and Running Color Channel Shift"
%% node8:::HeadingStyle
%% click node9 goToHeading "Max/Min Channel Filtering and Calculation"
%% node9:::HeadingStyle
%% click node10 goToHeading "Max/Min Channel Filtering and Calculation"
%% node10:::HeadingStyle
```

<SwmSnippet path="/Modules/Processor.bas" line="2030">

---

In <SwmToken path="Modules/Processor.bas" pos="2030:4:4" line-data="Private Function Process_AdjustmentsMenu(ByVal processID As String, Optional raiseDialog As Boolean = False, Optional processParameters As String = vbNullString, Optional createUndo As PD_UndoType = UNDO_Nothing, Optional relevantTool As Long = -1, Optional recordAction As Boolean = True, Optional ByRef returnDetails As String = vbNullString) As Boolean">`Process_AdjustmentsMenu`</SwmToken>, we match the <SwmToken path="Modules/Processor.bas" pos="2030:8:8" line-data="Private Function Process_AdjustmentsMenu(ByVal processID As String, Optional raiseDialog As Boolean = False, Optional processParameters As String = vbNullString, Optional createUndo As PD_UndoType = UNDO_Nothing, Optional relevantTool As Long = -1, Optional recordAction As Boolean = True, Optional ByRef returnDetails As String = vbNullString) As Boolean">`processID`</SwmToken> to a bunch of image adjustment actions—like brightness/contrast, curves, exposure, color balance, and more. If <SwmToken path="Modules/Processor.bas" pos="2030:17:17" line-data="Private Function Process_AdjustmentsMenu(ByVal processID As String, Optional raiseDialog As Boolean = False, Optional processParameters As String = vbNullString, Optional createUndo As PD_UndoType = UNDO_Nothing, Optional relevantTool As Long = -1, Optional recordAction As Boolean = True, Optional ByRef returnDetails As String = vbNullString) As Boolean">`raiseDialog`</SwmToken> is set, we call <SwmPath>[Modules/Interface.bas](Modules/Interface.bas)</SwmPath> to show the right dialog for user input; otherwise, we run the adjustment directly with the given parameters. This keeps all adjustment logic in one place and makes it easy to add new ones.

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

After <SwmPath>[Modules/Interface.bas](Modules/Interface.bas)</SwmPath> (for dialog-based adjustments), <SwmToken path="Modules/Processor.bas" pos="2130:1:1" line-data="        Process_AdjustmentsMenu = True">`Process_AdjustmentsMenu`</SwmToken> handles fixed adjustments like 'Film negative' by calling <SwmPath>[Modules/Filters_Color.bas](Modules/Filters_Color.bas)</SwmPath> directly. No dialog is needed, so the effect is applied right away.

```visual basic
    'Invert operations
    ElseIf Strings.StringsEqual(processID, "Film negative", True) Then
        MenuNegative
        Process_AdjustmentsMenu = True
    
    ElseIf Strings.StringsEqual(processID, "Invert hue", True) Then
```

---

</SwmSnippet>

### Film Negative Filter and Progress Updates

```mermaid
%%{init: {"flowchart": {"defaultRenderer": "elk"}} }%%
flowchart TD
    node1["Notify user: Calculating film negative values"]
    click node1 openCode "Modules/Filters_Color.bas:169:171"
    node1 --> node2["Prepare image data and progress bar"]
    click node2 openCode "Modules/Filters_Color.bas:173:193"
    
    subgraph loop1["For each row in selected area"]
        node2 --> node3["Invert colors of all pixels in row"]
        click node3 openCode "Modules/Filters_Color.bas:197:216"
        node3 --> node4["Update progress bar"]
        click node4 openCode "Modules/Filters_Color.bas:219:220"
        node4 --> node5{"User pressed ESC?"}
        click node5 openCode "Modules/Filters_Color.bas:218:218"
        node5 -->|"No"| node3
        node5 -->|"Yes"| node6["Exit loop and finalize"]
        click node6 openCode "Modules/Filters_Color.bas:218:229"
    end
    node6 --> node7["Finalize and display negative image"]
    click node7 openCode "Modules/Filters_Color.bas:223:229"
classDef HeadingStyle fill:#777777,stroke:#333,stroke-width:2px;

%% Swimm:
%% %%{init: {"flowchart": {"defaultRenderer": "elk"}} }%%
%% flowchart TD
%%     node1["Notify user: Calculating film negative values"]
%%     click node1 openCode "<SwmPath>[Modules/Filters_Color.bas](Modules/Filters_Color.bas)</SwmPath>:169:171"
%%     node1 --> node2["Prepare image data and progress bar"]
%%     click node2 openCode "<SwmPath>[Modules/Filters_Color.bas](Modules/Filters_Color.bas)</SwmPath>:173:193"
%%     
%%     subgraph loop1["For each row in selected area"]
%%         node2 --> node3["Invert colors of all pixels in row"]
%%         click node3 openCode "<SwmPath>[Modules/Filters_Color.bas](Modules/Filters_Color.bas)</SwmPath>:197:216"
%%         node3 --> node4["Update progress bar"]
%%         click node4 openCode "<SwmPath>[Modules/Filters_Color.bas](Modules/Filters_Color.bas)</SwmPath>:219:220"
%%         node4 --> node5{"User pressed ESC?"}
%%         click node5 openCode "<SwmPath>[Modules/Filters_Color.bas](Modules/Filters_Color.bas)</SwmPath>:218:218"
%%         node5 -->|"No"| node3
%%         node5 -->|"Yes"| node6["Exit loop and finalize"]
%%         click node6 openCode "<SwmPath>[Modules/Filters_Color.bas](Modules/Filters_Color.bas)</SwmPath>:218:229"
%%     end
%%     node6 --> node7["Finalize and display negative image"]
%%     click node7 openCode "<SwmPath>[Modules/Filters_Color.bas](Modules/Filters_Color.bas)</SwmPath>:223:229"
%% classDef HeadingStyle fill:#777777,stroke:#333,stroke-width:2px;
```

<SwmSnippet path="/Modules/Filters_Color.bas" line="169">

---

<SwmToken path="Modules/Filters_Color.bas" pos="169:4:4" line-data="Public Sub MenuNegative()">`MenuNegative`</SwmToken> shows a status message, preps the image data, and calls <SwmPath>[Modules/Interface.bas](Modules/Interface.bas)</SwmPath> to update the UI.

```visual basic
Public Sub MenuNegative()

    Message "Calculating film negative values..."

```

---

</SwmSnippet>

<SwmSnippet path="/Modules/Interface.bas" line="1668">

---

Message checks for duplicate messages and skips them to avoid UI spam. It also handles translation, parameter replacement, and logs messages if debug mode is on. If macros are recording, it adds a 'Recording' tag. Finally, it posts the message to the main canvas unless we're in batch mode.

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

After <SwmPath>[Modules/Interface.bas](Modules/Interface.bas)</SwmPath> updates the message, <SwmToken path="Modules/Processor.bas" pos="2129:1:1" line-data="        MenuNegative">`MenuNegative`</SwmToken> preps the image data for direct pixel access and runs the negative filter. It updates the progress bar only at calculated intervals to keep things fast, and checks for ESC to let the user cancel. Next, it calls <SwmPath>[Modules/ProgressBars.bas](Modules/ProgressBars.bas)</SwmPath> to update the UI progress bar and taskbar.

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

<SwmToken path="Modules/ProgressBars.bas" pos="43:4:4" line-data="Public Sub SetProgBarVal(ByVal pbVal As Double)">`SetProgBarVal`</SwmToken> updates the progress bar on the main canvas (unless we're in batch mode), updates the Windows taskbar if supported, and calls <SwmToken path="Modules/ProgressBars.bas" pos="54:3:3" line-data="        VBHacks.DoEvents_PaintOnly False">`DoEvents_PaintOnly`</SwmToken> to keep the UI responsive during long actions.

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

After the progress bar and pixel processing, we finalize the image changes and update the UI so the user sees the effect.

```visual basic
    'Safely deallocate imageData()
    workingDIB.UnwrapArrayFromDIB imageData
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering
    EffectPrep.FinalizeImageData
    
End Sub
```

---

</SwmSnippet>

### Handling Hue Inversion Command

```mermaid
%%{init: {"flowchart": {"defaultRenderer": "elk"}} }%%
flowchart TD
    node1["User selects an adjustment from the menu"]
    click node1 openCode "Modules/Processor.bas:2133:2136"
    node1 --> node2{"Which adjustment?"}
    click node2 openCode "Modules/Processor.bas:2133:2136"
    node2 -->|"Invert Hue"| node3["Apply hue inversion to image"]
    click node3 openCode "Modules/Processor.bas:2133:2134"
    node3 --> node5["Finish adjustment"]
    click node5 openCode "Modules/Processor.bas:2134:2134"
    node2 -->|"Invert RGB"| node4["Apply RGB inversion to image"]
    click node4 openCode "Modules/Processor.bas:2136:2136"
    node4 --> node5
classDef HeadingStyle fill:#777777,stroke:#333,stroke-width:2px;

%% Swimm:
%% %%{init: {"flowchart": {"defaultRenderer": "elk"}} }%%
%% flowchart TD
%%     node1["User selects an adjustment from the menu"]
%%     click node1 openCode "<SwmPath>[Modules/Processor.bas](Modules/Processor.bas)</SwmPath>:2133:2136"
%%     node1 --> node2{"Which adjustment?"}
%%     click node2 openCode "<SwmPath>[Modules/Processor.bas](Modules/Processor.bas)</SwmPath>:2133:2136"
%%     node2 -->|"Invert Hue"| node3["Apply hue inversion to image"]
%%     click node3 openCode "<SwmPath>[Modules/Processor.bas](Modules/Processor.bas)</SwmPath>:2133:2134"
%%     node3 --> node5["Finish adjustment"]
%%     click node5 openCode "<SwmPath>[Modules/Processor.bas](Modules/Processor.bas)</SwmPath>:2134:2134"
%%     node2 -->|"Invert RGB"| node4["Apply RGB inversion to image"]
%%     click node4 openCode "<SwmPath>[Modules/Processor.bas](Modules/Processor.bas)</SwmPath>:2136:2136"
%%     node4 --> node5
%% classDef HeadingStyle fill:#777777,stroke:#333,stroke-width:2px;
```

<SwmSnippet path="/Modules/Processor.bas" line="2133">

---

Back in <SwmToken path="Modules/Processor.bas" pos="2134:1:1" line-data="        Process_AdjustmentsMenu = True">`Process_AdjustmentsMenu`</SwmToken>, after <SwmToken path="Modules/Processor.bas" pos="2129:1:1" line-data="        MenuNegative">`MenuNegative`</SwmToken> finishes, we check if the next <SwmToken path="Modules/Processor.bas" pos="2136:7:7" line-data="    ElseIf Strings.StringsEqual(processID, &quot;Invert RGB&quot;, True) Then">`processID`</SwmToken> matches another filter (like 'Invert Hue') and call the corresponding function in <SwmPath>[Modules/Filters_Color.bas](Modules/Filters_Color.bas)</SwmPath>. This keeps the adjustment menu responsive to whatever action the user picked.

```visual basic
        MenuInvertHue
        Process_AdjustmentsMenu = True
        
    ElseIf Strings.StringsEqual(processID, "Invert RGB", True) Then
```

---

</SwmSnippet>

### Dispatching and Running Hue Inversion

```mermaid
%%{init: {"flowchart": {"defaultRenderer": "elk"}} }%%
flowchart TD
    node1["Start hue inversion for selected area"]
    click node1 openCode "Modules/Filters_Color.bas:232:234"
    
    subgraph loop1["For each row in selected area"]
        node2["For each pixel in row"]
        click node2 openCode "Modules/Filters_Color.bas:260:283"
        subgraph loop2["For each pixel in row"]
            node3["Convert pixel color to HSL"]
            click node3 openCode "Modules/Filters_Color.bas:263:270"
            node4["Invert hue value"]
            click node4 openCode "Modules/Filters_Color.bas:273:273"
            node5["Convert color back to RGB"]
            click node5 openCode "Modules/Filters_Color.bas:275:276"
            node6["Update pixel color"]
            click node6 openCode "Modules/Filters_Color.bas:279:281"
        end
        node7{"User pressed ESC?"}
        click node7 openCode "Modules/Filters_Color.bas:285:285"
        node2 --> node7
        node7 -->|"No"| node2
        node7 -->|"Yes"| node8["Stop processing"]
        click node8 openCode "Modules/Filters_Color.bas:285:285"
    end
    node1 --> node2
    node2 --> node9["Finalize and update image"]
    click node9 openCode "Modules/Filters_Color.bas:290:295"
classDef HeadingStyle fill:#777777,stroke:#333,stroke-width:2px;

%% Swimm:
%% %%{init: {"flowchart": {"defaultRenderer": "elk"}} }%%
%% flowchart TD
%%     node1["Start hue inversion for selected area"]
%%     click node1 openCode "<SwmPath>[Modules/Filters_Color.bas](Modules/Filters_Color.bas)</SwmPath>:232:234"
%%     
%%     subgraph loop1["For each row in selected area"]
%%         node2["For each pixel in row"]
%%         click node2 openCode "<SwmPath>[Modules/Filters_Color.bas](Modules/Filters_Color.bas)</SwmPath>:260:283"
%%         subgraph loop2["For each pixel in row"]
%%             node3["Convert pixel color to HSL"]
%%             click node3 openCode "<SwmPath>[Modules/Filters_Color.bas](Modules/Filters_Color.bas)</SwmPath>:263:270"
%%             node4["Invert hue value"]
%%             click node4 openCode "<SwmPath>[Modules/Filters_Color.bas](Modules/Filters_Color.bas)</SwmPath>:273:273"
%%             node5["Convert color back to RGB"]
%%             click node5 openCode "<SwmPath>[Modules/Filters_Color.bas](Modules/Filters_Color.bas)</SwmPath>:275:276"
%%             node6["Update pixel color"]
%%             click node6 openCode "<SwmPath>[Modules/Filters_Color.bas](Modules/Filters_Color.bas)</SwmPath>:279:281"
%%         end
%%         node7{"User pressed ESC?"}
%%         click node7 openCode "<SwmPath>[Modules/Filters_Color.bas](Modules/Filters_Color.bas)</SwmPath>:285:285"
%%         node2 --> node7
%%         node7 -->|"No"| node2
%%         node7 -->|"Yes"| node8["Stop processing"]
%%         click node8 openCode "<SwmPath>[Modules/Filters_Color.bas](Modules/Filters_Color.bas)</SwmPath>:285:285"
%%     end
%%     node1 --> node2
%%     node2 --> node9["Finalize and update image"]
%%     click node9 openCode "<SwmPath>[Modules/Filters_Color.bas](Modules/Filters_Color.bas)</SwmPath>:290:295"
%% classDef HeadingStyle fill:#777777,stroke:#333,stroke-width:2px;
```

<SwmSnippet path="/Modules/Filters_Color.bas" line="232">

---

We show a status message so the user knows the effect is running.

```visual basic
Public Sub MenuInvertHue()

    Message "Inverting..."

```

---

</SwmSnippet>

<SwmSnippet path="/Modules/Filters_Color.bas" line="236">

---

After showing the status message, <SwmToken path="Modules/Processor.bas" pos="2133:1:1" line-data="        MenuInvertHue">`MenuInvertHue`</SwmToken> grabs the image pixel data into a byte array, loops over each pixel, converts RGB to HSL, inverts the hue, and writes the new RGB values back. <SwmToken path="Modules/Filters_Color.bas" pos="250:1:1" line-data="    ProgressBars.SetProgBarMax finalY">`ProgressBars`</SwmToken> is used to update the UI only at certain intervals, and we check for ESC to let the user cancel. This keeps things fast and responsive.

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

After the progress bar and pixel loop, <SwmToken path="Modules/Processor.bas" pos="2133:1:1" line-data="        MenuInvertHue">`MenuInvertHue`</SwmToken> finishes by unwrapping the image data and calling <SwmToken path="Modules/Filters_Color.bas" pos="294:1:3" line-data="    EffectPrep.FinalizeImageData">`EffectPrep.FinalizeImageData`</SwmToken>. This step commits the changes and updates the UI so the user sees the result.

```visual basic
    'Safely deallocate imageData()
    workingDIB.UnwrapArrayFromDIB imageData
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering
    EffectPrep.FinalizeImageData
    
End Sub
```

---

</SwmSnippet>

### Handling RGB Inversion Command

<SwmSnippet path="/Modules/Processor.bas" line="2137">

---

Back in <SwmToken path="Modules/Processor.bas" pos="2138:1:1" line-data="        Process_AdjustmentsMenu = True">`Process_AdjustmentsMenu`</SwmToken>, after <SwmToken path="Modules/Processor.bas" pos="2133:1:1" line-data="        MenuInvertHue">`MenuInvertHue`</SwmToken>, we check if the <SwmToken path="Modules/Processor.bas" pos="89:8:8" line-data="Public Sub Process(ByVal processID As String, Optional raiseDialog As Boolean = False, Optional processParameters As String = vbNullString, Optional createUndo As PD_UndoType = UNDO_Nothing, Optional relevantTool As Long = -1, Optional recordAction As Boolean = True)">`processID`</SwmToken> matches 'Invert RGB' and call <SwmToken path="Modules/Processor.bas" pos="2137:1:1" line-data="        MenuInvert">`MenuInvert`</SwmToken> in <SwmPath>[Modules/Filters_Color.bas](Modules/Filters_Color.bas)</SwmPath> if so. This keeps the adjustment menu handling each action separately.

```visual basic
        MenuInvert
        Process_AdjustmentsMenu = True
    
```

---

</SwmSnippet>

### Dispatching and Running RGB Inversion

```mermaid
%%{init: {"flowchart": {"defaultRenderer": "elk"}} }%%
flowchart TD
    node1["Start color inversion"]
    click node1 openCode "Modules/Filters_Color.bas:57:59"
    node1 --> node2["Determine selected area"]
    click node2 openCode "Modules/Filters_Color.bas:66:69"
    node2 --> node3["Prepare for inversion"]
    click node3 openCode "Modules/Filters_Color.bas:61:85"
    
    subgraph loop1["For each row in selected area"]
      node3 --> node4["For each pixel in row, invert color"]
      click node4 openCode "Modules/Filters_Color.bas:87:93"
      node4 --> node5{"User pressed ESC?"}
      click node5 openCode "Modules/Filters_Color.bas:94:95"
      node5 -->|"Yes"| node6["Stop inversion"]
      click node6 openCode "Modules/Filters_Color.bas:95:95"
      node5 -->|"No"| node8["Continue to next row"]
      click node8 openCode "Modules/Filters_Color.bas:98:98"
    end
    loop1 --> node7["Finalize inversion and update image"]
    click node7 openCode "Modules/Filters_Color.bas:100:104"
classDef HeadingStyle fill:#777777,stroke:#333,stroke-width:2px;

%% Swimm:
%% %%{init: {"flowchart": {"defaultRenderer": "elk"}} }%%
%% flowchart TD
%%     node1["Start color inversion"]
%%     click node1 openCode "<SwmPath>[Modules/Filters_Color.bas](Modules/Filters_Color.bas)</SwmPath>:57:59"
%%     node1 --> node2["Determine selected area"]
%%     click node2 openCode "<SwmPath>[Modules/Filters_Color.bas](Modules/Filters_Color.bas)</SwmPath>:66:69"
%%     node2 --> node3["Prepare for inversion"]
%%     click node3 openCode "<SwmPath>[Modules/Filters_Color.bas](Modules/Filters_Color.bas)</SwmPath>:61:85"
%%     
%%     subgraph loop1["For each row in selected area"]
%%       node3 --> node4["For each pixel in row, invert color"]
%%       click node4 openCode "<SwmPath>[Modules/Filters_Color.bas](Modules/Filters_Color.bas)</SwmPath>:87:93"
%%       node4 --> node5{"User pressed ESC?"}
%%       click node5 openCode "<SwmPath>[Modules/Filters_Color.bas](Modules/Filters_Color.bas)</SwmPath>:94:95"
%%       node5 -->|"Yes"| node6["Stop inversion"]
%%       click node6 openCode "<SwmPath>[Modules/Filters_Color.bas](Modules/Filters_Color.bas)</SwmPath>:95:95"
%%       node5 -->|"No"| node8["Continue to next row"]
%%       click node8 openCode "<SwmPath>[Modules/Filters_Color.bas](Modules/Filters_Color.bas)</SwmPath>:98:98"
%%     end
%%     loop1 --> node7["Finalize inversion and update image"]
%%     click node7 openCode "<SwmPath>[Modules/Filters_Color.bas](Modules/Filters_Color.bas)</SwmPath>:100:104"
%% classDef HeadingStyle fill:#777777,stroke:#333,stroke-width:2px;
```

<SwmSnippet path="/Modules/Filters_Color.bas" line="57">

---

We show a status message so the user knows the effect is running.

```visual basic
Public Sub MenuInvert()
        
    Message "Inverting..."
    
```

---

</SwmSnippet>

<SwmSnippet path="/Modules/Filters_Color.bas" line="61">

---

We invert each pixel's RGB, update the progress bar, and finalize the image for display.

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

### Handling Color Mapping and Monochrome Commands

```mermaid
%%{init: {"flowchart": {"defaultRenderer": "elk"}} }%%
flowchart TD
    node1["User selects a color adjustment or mapping"] --> node2{"Which adjustment? (processID)"}
    click node1 openCode "Modules/Processor.bas:2140:2180"
    click node2 openCode "Modules/Processor.bas:2141:2180"
    node2 -->|"Gradient map, Palette map, Color to monochrome, Channel mixer, Rechannel"| node3{"Show dialog for user input? (raiseDialog)"}
    node2 -->|"Shift colors (left/right)"| node4["Apply color shift to image"]
    node2 -->|"Maximum/Minimum channel"| node5["Apply max/min channel effect"]
    click node3 openCode "Modules/Processor.bas:2142:2166"
    click node4 openCode "Modules/Processor.bas:2168:2174"
    click node5 openCode "Modules/Processor.bas:2176:2180"
    node3 -->|"Yes"| node6["Show adjustment dialog to user"]
    node3 -->|"No"| node7["Apply adjustment effect immediately"]
    click node6 openCode "Modules/Processor.bas:2142:2166"
    click node7 openCode "Modules/Processor.bas:2142:2166"
    node4 --> node8["Operation complete"]
    node5 --> node8
    node6 --> node8
    node7 --> node8
    click node8 openCode "Modules/Processor.bas:2143:2180"
classDef HeadingStyle fill:#777777,stroke:#333,stroke-width:2px;

%% Swimm:
%% %%{init: {"flowchart": {"defaultRenderer": "elk"}} }%%
%% flowchart TD
%%     node1["User selects a color adjustment or mapping"] --> node2{"Which adjustment? (<SwmToken path="Modules/Processor.bas" pos="89:8:8" line-data="Public Sub Process(ByVal processID As String, Optional raiseDialog As Boolean = False, Optional processParameters As String = vbNullString, Optional createUndo As PD_UndoType = UNDO_Nothing, Optional relevantTool As Long = -1, Optional recordAction As Boolean = True)">`processID`</SwmToken>)"}
%%     click node1 openCode "<SwmPath>[Modules/Processor.bas](Modules/Processor.bas)</SwmPath>:2140:2180"
%%     click node2 openCode "<SwmPath>[Modules/Processor.bas](Modules/Processor.bas)</SwmPath>:2141:2180"
%%     node2 -->|"Gradient map, Palette map, Color to monochrome, Channel mixer, Rechannel"| node3{"Show dialog for user input? (<SwmToken path="Modules/Processor.bas" pos="89:17:17" line-data="Public Sub Process(ByVal processID As String, Optional raiseDialog As Boolean = False, Optional processParameters As String = vbNullString, Optional createUndo As PD_UndoType = UNDO_Nothing, Optional relevantTool As Long = -1, Optional recordAction As Boolean = True)">`raiseDialog`</SwmToken>)"}
%%     node2 -->|"Shift colors (left/right)"| node4["Apply color shift to image"]
%%     node2 -->|"Maximum/Minimum channel"| node5["Apply <SwmToken path="Modules/Interface.bas" pos="900:23:25" line-data="            &#39; I&#39;m currently using the &quot;rule of three&quot; - max/min values are the current dimensions of the image, x3.">`max/min`</SwmToken> channel effect"]
%%     click node3 openCode "<SwmPath>[Modules/Processor.bas](Modules/Processor.bas)</SwmPath>:2142:2166"
%%     click node4 openCode "<SwmPath>[Modules/Processor.bas](Modules/Processor.bas)</SwmPath>:2168:2174"
%%     click node5 openCode "<SwmPath>[Modules/Processor.bas](Modules/Processor.bas)</SwmPath>:2176:2180"
%%     node3 -->|"Yes"| node6["Show adjustment dialog to user"]
%%     node3 -->|"No"| node7["Apply adjustment effect immediately"]
%%     click node6 openCode "<SwmPath>[Modules/Processor.bas](Modules/Processor.bas)</SwmPath>:2142:2166"
%%     click node7 openCode "<SwmPath>[Modules/Processor.bas](Modules/Processor.bas)</SwmPath>:2142:2166"
%%     node4 --> node8["Operation complete"]
%%     node5 --> node8
%%     node6 --> node8
%%     node7 --> node8
%%     click node8 openCode "<SwmPath>[Modules/Processor.bas](Modules/Processor.bas)</SwmPath>:2143:2180"
%% classDef HeadingStyle fill:#777777,stroke:#333,stroke-width:2px;
```

<SwmSnippet path="/Modules/Processor.bas" line="2140">

---

Back in <SwmToken path="Modules/Processor.bas" pos="2143:1:1" line-data="        Process_AdjustmentsMenu = True">`Process_AdjustmentsMenu`</SwmToken>, after color inversion, we check if the next <SwmToken path="Modules/Processor.bas" pos="2141:7:7" line-data="    ElseIf Strings.StringsEqual(processID, &quot;Gradient map&quot;, True) Then">`processID`</SwmToken> matches a mapping or monochrome command. If <SwmToken path="Modules/Processor.bas" pos="2142:3:3" line-data="        If raiseDialog Then ShowPDDialog vbModal, FormGradientMap Else FormGradientMap.ApplyGradientMap processParameters">`raiseDialog`</SwmToken> is set, we show the relevant dialog via <SwmPath>[Modules/Interface.bas](Modules/Interface.bas)</SwmPath>; otherwise, we run the effect directly. This keeps the UI flexible for both interactive and automated use.

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

Back in <SwmToken path="Modules/Processor.bas" pos="2174:1:1" line-data="        Process_AdjustmentsMenu = True">`Process_AdjustmentsMenu`</SwmToken>, after handling dialogs, we check if the <SwmToken path="Modules/Processor.bas" pos="2176:7:7" line-data="    ElseIf Strings.StringsEqual(processID, &quot;Maximum channel&quot;, True) Then">`processID`</SwmToken> matches a color shift command. If so, we call <SwmToken path="Modules/Processor.bas" pos="2173:1:1" line-data="        MenuCShift False">`MenuCShift`</SwmToken> in <SwmPath>[Modules/Filters_Color.bas](Modules/Filters_Color.bas)</SwmPath> with the right direction. This keeps the adjustment menu handling each color operation as requested.

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

### Dispatching and Running Color Channel Shift

```mermaid
%%{init: {"flowchart": {"defaultRenderer": "elk"}} }%%
flowchart TD
    node1["User initiates color shift"]
    click node1 openCode "Modules/Filters_Color.bas:109:111"
    node1 --> node2["Prepare image data"]
    click node2 openCode "Modules/Filters_Color.bas:113:133"
    
    subgraph loop1["For each pixel in selected area"]
        node2 --> node3{"Shift left?"}
        click node3 openCode "Modules/Filters_Color.bas:139:147"
        node3 -->|"Yes"| node4["Assign new color channels (left shift: G→R, R→B, B→G)"]
        click node4 openCode "Modules/Filters_Color.bas:140:143"
        node3 -->|"No"| node5["Assign new color channels (right shift: R→B, B→G, G→R)"]
        click node5 openCode "Modules/Filters_Color.bas:144:147"
    end
    loop1 --> node6["Finalize and display result"]
    click node6 openCode "Modules/Filters_Color.bas:160:166"

classDef HeadingStyle fill:#777777,stroke:#333,stroke-width:2px;

%% Swimm:
%% %%{init: {"flowchart": {"defaultRenderer": "elk"}} }%%
%% flowchart TD
%%     node1["User initiates color shift"]
%%     click node1 openCode "<SwmPath>[Modules/Filters_Color.bas](Modules/Filters_Color.bas)</SwmPath>:109:111"
%%     node1 --> node2["Prepare image data"]
%%     click node2 openCode "<SwmPath>[Modules/Filters_Color.bas](Modules/Filters_Color.bas)</SwmPath>:113:133"
%%     
%%     subgraph loop1["For each pixel in selected area"]
%%         node2 --> node3{"Shift left?"}
%%         click node3 openCode "<SwmPath>[Modules/Filters_Color.bas](Modules/Filters_Color.bas)</SwmPath>:139:147"
%%         node3 -->|"Yes"| node4["Assign new color channels (left shift: G→R, R→B, B→G)"]
%%         click node4 openCode "<SwmPath>[Modules/Filters_Color.bas](Modules/Filters_Color.bas)</SwmPath>:140:143"
%%         node3 -->|"No"| node5["Assign new color channels (right shift: R→B, B→G, G→R)"]
%%         click node5 openCode "<SwmPath>[Modules/Filters_Color.bas](Modules/Filters_Color.bas)</SwmPath>:144:147"
%%     end
%%     loop1 --> node6["Finalize and display result"]
%%     click node6 openCode "<SwmPath>[Modules/Filters_Color.bas](Modules/Filters_Color.bas)</SwmPath>:160:166"
%% 
%% classDef HeadingStyle fill:#777777,stroke:#333,stroke-width:2px;
```

<SwmSnippet path="/Modules/Filters_Color.bas" line="109">

---

We show a status message so the user knows the effect is running.

```visual basic
Public Sub MenuCShift(Optional ByVal shiftLeft As Boolean = False)
    
    Message "Shifting RGB values..."
    
```

---

</SwmSnippet>

<SwmSnippet path="/Modules/Filters_Color.bas" line="113">

---

We shift the RGB channels for each pixel, update the progress bar at intervals, and finalize the image for display.

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

After the progress bar and pixel loop, <SwmToken path="Modules/Processor.bas" pos="2169:1:1" line-data="        MenuCShift True">`MenuCShift`</SwmToken> finishes by unwrapping the image data and calling <SwmToken path="Modules/Filters_Color.bas" pos="164:1:3" line-data="    EffectPrep.FinalizeImageData">`EffectPrep.FinalizeImageData`</SwmToken>. This step commits the changes and updates the UI so the user sees the result.

```visual basic
    'Safely deallocate imageData()
    workingDIB.UnwrapArrayFromDIB imageData
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering
    EffectPrep.FinalizeImageData
    
End Sub
```

---

</SwmSnippet>

### Channel Isolation Adjustment Dispatch

<SwmSnippet path="/Modules/Processor.bas" line="2181">

---

After returning from <SwmToken path="Modules/Processor.bas" pos="2169:1:1" line-data="        MenuCShift True">`MenuCShift`</SwmToken> in <SwmPath>[Modules/Filters_Color.bas](Modules/Filters_Color.bas)</SwmPath>, <SwmToken path="Modules/Processor.bas" pos="2182:1:1" line-data="        Process_AdjustmentsMenu = True">`Process_AdjustmentsMenu`</SwmToken> calls <SwmToken path="Modules/Processor.bas" pos="2181:1:1" line-data="        FilterMaxMinChannel False">`FilterMaxMinChannel`</SwmToken> to isolate either the max or min color channel per pixel. This keeps the adjustment chain moving, and sets <SwmToken path="Modules/Processor.bas" pos="2182:1:1" line-data="        Process_AdjustmentsMenu = True">`Process_AdjustmentsMenu`</SwmToken> = True to mark the operation as handled.

```visual basic
        FilterMaxMinChannel False
        Process_AdjustmentsMenu = True
        
```

---

</SwmSnippet>

### Max/Min Channel Filtering and Calculation

```mermaid
%%{init: {"flowchart": {"defaultRenderer": "elk"}} }%%
flowchart TD
    node1["User chooses to isolate max or min color channel"]
    click node1 openCode "Modules/Filters_Color.bas:299:305"
    node1 --> node2{"Isolate maximum channel?"}
    click node2 openCode "Modules/Filters_Color.bas:301:305"
    
    subgraph loop1["For each pixel in selected area"]
        node2 -->|"Yes"| node3["Find maximum of red, green, blue"]
        click node3 openCode "Modules/Filters_Color.bas:335:336"
        node3 --> node4["Set non-maximum channels to zero"]
        click node4 openCode "Modules/Filters_Color.bas:337:340"
        node2 -->|"No"| node5["Find minimum of red, green, blue"]
        click node5 openCode "Modules/Filters_Color.bas:341"
        node5 --> node6["Set non-minimum channels to zero"]
        click node6 openCode "Modules/Filters_Color.bas:342:344"
    end
    node4 --> node7["Finalize image"]
    node6 --> node7
    click node7 openCode "Modules/Filters_Color.bas:358:364"
classDef HeadingStyle fill:#777777,stroke:#333,stroke-width:2px;

%% Swimm:
%% %%{init: {"flowchart": {"defaultRenderer": "elk"}} }%%
%% flowchart TD
%%     node1["User chooses to isolate max or min color channel"]
%%     click node1 openCode "<SwmPath>[Modules/Filters_Color.bas](Modules/Filters_Color.bas)</SwmPath>:299:305"
%%     node1 --> node2{"Isolate maximum channel?"}
%%     click node2 openCode "<SwmPath>[Modules/Filters_Color.bas](Modules/Filters_Color.bas)</SwmPath>:301:305"
%%     
%%     subgraph loop1["For each pixel in selected area"]
%%         node2 -->|"Yes"| node3["Find maximum of red, green, blue"]
%%         click node3 openCode "<SwmPath>[Modules/Filters_Color.bas](Modules/Filters_Color.bas)</SwmPath>:335:336"
%%         node3 --> node4["Set non-maximum channels to zero"]
%%         click node4 openCode "<SwmPath>[Modules/Filters_Color.bas](Modules/Filters_Color.bas)</SwmPath>:337:340"
%%         node2 -->|"No"| node5["Find minimum of red, green, blue"]
%%         click node5 openCode "<SwmPath>[Modules/Filters_Color.bas](Modules/Filters_Color.bas)</SwmPath>:341"
%%         node5 --> node6["Set non-minimum channels to zero"]
%%         click node6 openCode "<SwmPath>[Modules/Filters_Color.bas](Modules/Filters_Color.bas)</SwmPath>:342:344"
%%     end
%%     node4 --> node7["Finalize image"]
%%     node6 --> node7
%%     click node7 openCode "<SwmPath>[Modules/Filters_Color.bas](Modules/Filters_Color.bas)</SwmPath>:358:364"
%% classDef HeadingStyle fill:#777777,stroke:#333,stroke-width:2px;
```

<SwmSnippet path="/Modules/Filters_Color.bas" line="299">

---

In <SwmToken path="Modules/Filters_Color.bas" pos="299:4:4" line-data="Public Sub FilterMaxMinChannel(ByVal useMax As Boolean)">`FilterMaxMinChannel`</SwmToken>, we start by showing a message to let the user know if we're isolating max or min channels. This sets up the UI feedback before we touch any pixel data.

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

After showing the message, <SwmToken path="Modules/Processor.bas" pos="2177:1:1" line-data="        FilterMaxMinChannel True">`FilterMaxMinChannel`</SwmToken> grabs the image pixel data and loops over each pixel. For each pixel, it calls <SwmToken path="Modules/Filters_Color.bas" pos="336:5:5" line-data="            maxVal = Max3Int(r, g, b)">`Max3Int`</SwmToken> or <SwmToken path="Modules/Filters_Color.bas" pos="341:5:5" line-data="            minVal = Min3Int(r, g, b)">`Min3Int`</SwmToken> to figure out which channel to keep, zeroes out the rest, and writes the result back. Progress bar updates and ESC checks keep things responsive.

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

<SwmToken path="Modules/PDMath.bas" pos="607:4:4" line-data="Public Function Max3Int(ByVal rR As Long, ByVal rG As Long, ByVal rB As Long) As Long">`Max3Int`</SwmToken> just picks the largest of three color channel values using nested Ifs. It's called from the pixel loop to isolate the max channel per pixel.

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

After using <SwmToken path="Modules/Filters_Color.bas" pos="336:5:5" line-data="            maxVal = Max3Int(r, g, b)">`Max3Int`</SwmToken> for max channel isolation, we call <SwmToken path="Modules/Filters_Color.bas" pos="341:5:5" line-data="            minVal = Min3Int(r, g, b)">`Min3Int`</SwmToken> for min channel isolation if <SwmToken path="Modules/Filters_Color.bas" pos="299:8:8" line-data="Public Sub FilterMaxMinChannel(ByVal useMax As Boolean)">`useMax`</SwmToken> is False. The pixel loop zeroes out channels above the min, then writes the result back.

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

<SwmToken path="Modules/PDMath.bas" pos="633:4:4" line-data="Public Function Min3Int(ByVal rR As Long, ByVal rG As Long, ByVal rB As Long) As Long">`Min3Int`</SwmToken> returns the smallest of three color channel values using nested Ifs. It's called from the pixel loop for min channel isolation.

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

We update the progress bar and check for ESC only when needed to keep things fast.

```visual basic
            SetProgBarVal y
        End If
    Next y
        
```

---

</SwmSnippet>

<SwmSnippet path="/Modules/Filters_Color.bas" line="358">

---

After the pixel loop and progress bar updates, we unwrap the image data and call <SwmToken path="Modules/Filters_Color.bas" pos="362:3:3" line-data="    EffectPrep.FinalizeImageData">`FinalizeImageData`</SwmToken> to commit the changes and refresh the UI.

```visual basic
    'Safely deallocate imageData()
    workingDIB.UnwrapArrayFromDIB imageData
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering
    EffectPrep.FinalizeImageData
    
End Sub
```

---

</SwmSnippet>

### Histogram Adjustment and Finalization

```mermaid
%%{init: {"flowchart": {"defaultRenderer": "elk"}} }%%
flowchart TD
    node1["User selects an adjustment from the menu"]
    click node1 openCode "Modules/Processor.bas:2184:2199"
    node1 --> node2{"processID?"}
    click node2 openCode "Modules/Processor.bas:2185:2197"
    node2 -->|"Display histogram"| node3["Show histogram dialog"]
    click node3 openCode "Modules/Processor.bas:2186:2187"
    node2 -->|"Stretch histogram"| node4["Stretch the histogram"]
    click node4 openCode "Modules/Processor.bas:2190:2191"
    node2 -->|"Equalize"| node5{"raiseDialog?"}
    click node5 openCode "Modules/Processor.bas:2194:2195"
    node5 -->|"Yes"| node6["Show equalize dialog"]
    click node6 openCode "Modules/Processor.bas:2194:2195"
    node5 -->|"No"| node7["Equalize histogram directly"]
    click node7 openCode "Modules/Processor.bas:2194:2195"
classDef HeadingStyle fill:#777777,stroke:#333,stroke-width:2px;

%% Swimm:
%% %%{init: {"flowchart": {"defaultRenderer": "elk"}} }%%
%% flowchart TD
%%     node1["User selects an adjustment from the menu"]
%%     click node1 openCode "<SwmPath>[Modules/Processor.bas](Modules/Processor.bas)</SwmPath>:2184:2199"
%%     node1 --> node2{"<SwmToken path="Modules/Processor.bas" pos="89:8:8" line-data="Public Sub Process(ByVal processID As String, Optional raiseDialog As Boolean = False, Optional processParameters As String = vbNullString, Optional createUndo As PD_UndoType = UNDO_Nothing, Optional relevantTool As Long = -1, Optional recordAction As Boolean = True)">`processID`</SwmToken>?"}
%%     click node2 openCode "<SwmPath>[Modules/Processor.bas](Modules/Processor.bas)</SwmPath>:2185:2197"
%%     node2 -->|"Display histogram"| node3["Show histogram dialog"]
%%     click node3 openCode "<SwmPath>[Modules/Processor.bas](Modules/Processor.bas)</SwmPath>:2186:2187"
%%     node2 -->|"Stretch histogram"| node4["Stretch the histogram"]
%%     click node4 openCode "<SwmPath>[Modules/Processor.bas](Modules/Processor.bas)</SwmPath>:2190:2191"
%%     node2 -->|"Equalize"| node5{"<SwmToken path="Modules/Processor.bas" pos="89:17:17" line-data="Public Sub Process(ByVal processID As String, Optional raiseDialog As Boolean = False, Optional processParameters As String = vbNullString, Optional createUndo As PD_UndoType = UNDO_Nothing, Optional relevantTool As Long = -1, Optional recordAction As Boolean = True)">`raiseDialog`</SwmToken>?"}
%%     click node5 openCode "<SwmPath>[Modules/Processor.bas](Modules/Processor.bas)</SwmPath>:2194:2195"
%%     node5 -->|"Yes"| node6["Show equalize dialog"]
%%     click node6 openCode "<SwmPath>[Modules/Processor.bas](Modules/Processor.bas)</SwmPath>:2194:2195"
%%     node5 -->|"No"| node7["Equalize histogram directly"]
%%     click node7 openCode "<SwmPath>[Modules/Processor.bas](Modules/Processor.bas)</SwmPath>:2194:2195"
%% classDef HeadingStyle fill:#777777,stroke:#333,stroke-width:2px;
```

<SwmSnippet path="/Modules/Processor.bas" line="2184">

---

After <SwmToken path="Modules/Processor.bas" pos="2177:1:1" line-data="        FilterMaxMinChannel True">`FilterMaxMinChannel`</SwmToken>, <SwmToken path="Modules/Processor.bas" pos="2187:1:1" line-data="        Process_AdjustmentsMenu = True">`Process_AdjustmentsMenu`</SwmToken> checks for histogram-related <SwmToken path="Modules/Processor.bas" pos="209:15:15" line-data="    &#39;For ease of reference, the various processIDs are divided into categories of similar functions.  These categories">`processIDs`</SwmToken>. For 'Display histogram', we show the dialog; for 'Stretch histogram' and 'Equalize', we run the effect or show the dialog based on <SwmToken path="Modules/Processor.bas" pos="2194:3:3" line-data="        If raiseDialog Then ShowPDDialog vbModal, FormEqualize Else FormEqualize.EqualizeHistogram processParameters">`raiseDialog`</SwmToken>. Each handled adjustment sets the return value to True.

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

## Effects Menu Dispatch

<SwmSnippet path="/Modules/Processor.bas" line="255">

---

After <SwmToken path="Modules/Processor.bas" pos="253:15:15" line-data="    If (Not processFound) Then processFound = Process_AdjustmentsMenu(processID, raiseDialog, processParameters, createUndo, relevantTool, recordAction, returnDetails)">`Process_AdjustmentsMenu`</SwmToken>, Process checks if <SwmToken path="Modules/Processor.bas" pos="256:6:6" line-data="    If (Not processFound) Then processFound = Process_EffectsMenu(processID, raiseDialog, processParameters, createUndo, relevantTool, recordAction, returnDetails)">`processFound`</SwmToken> is still False. If so, we call <SwmToken path="Modules/Processor.bas" pos="256:15:15" line-data="    If (Not processFound) Then processFound = Process_EffectsMenu(processID, raiseDialog, processParameters, createUndo, relevantTool, recordAction, returnDetails)">`Process_EffectsMenu`</SwmToken> to handle any remaining effect operations based on <SwmToken path="Modules/Processor.bas" pos="256:17:17" line-data="    If (Not processFound) Then processFound = Process_EffectsMenu(processID, raiseDialog, processParameters, createUndo, relevantTool, recordAction, returnDetails)">`processID`</SwmToken>.

```visual basic
    'Effects menu operations have been successfully migrated to XML strings.  (None of their functions raise special return conditions, FYI.)
    If (Not processFound) Then processFound = Process_EffectsMenu(processID, raiseDialog, processParameters, createUndo, relevantTool, recordAction, returnDetails)
    
```

---

</SwmSnippet>

## Effect Processing and Dialog Routing

<SwmSnippet path="/Modules/Processor.bas" line="1640">

---

In <SwmToken path="Modules/Processor.bas" pos="1640:4:4" line-data="Private Function Process_EffectsMenu(ByVal processID As String, Optional raiseDialog As Boolean = False, Optional processParameters As String = vbNullString, Optional createUndo As PD_UndoType = UNDO_Nothing, Optional relevantTool As Long = -1, Optional recordAction As Boolean = True, Optional ByRef returnDetails As String = vbNullString) As Boolean">`Process_EffectsMenu`</SwmToken>, we match the <SwmToken path="Modules/Processor.bas" pos="1640:8:8" line-data="Private Function Process_EffectsMenu(ByVal processID As String, Optional raiseDialog As Boolean = False, Optional processParameters As String = vbNullString, Optional createUndo As PD_UndoType = UNDO_Nothing, Optional relevantTool As Long = -1, Optional recordAction As Boolean = True, Optional ByRef returnDetails As String = vbNullString) As Boolean">`processID`</SwmToken> to known effects. For each, we either show a dialog (if <SwmToken path="Modules/Processor.bas" pos="1640:17:17" line-data="Private Function Process_EffectsMenu(ByVal processID As String, Optional raiseDialog As Boolean = False, Optional processParameters As String = vbNullString, Optional createUndo As PD_UndoType = UNDO_Nothing, Optional relevantTool As Long = -1, Optional recordAction As Boolean = True, Optional ByRef returnDetails As String = vbNullString) As Boolean">`raiseDialog`</SwmToken> is set) or run the effect directly. This keeps effect handling centralized and easy to extend.

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

After handling dialog-based effects, <SwmToken path="Modules/Processor.bas" pos="1721:1:1" line-data="        Process_EffectsMenu = True">`Process_EffectsMenu`</SwmToken> checks for legacy or <SwmToken path="Modules/Processor.bas" pos="1187:15:17" line-data="        &#39;If we need to do a &quot;special-case whole-image&quot; rasterization, do so now.">`special-case`</SwmToken> effects like Grid blur. These are called directly (<SwmToken path="Modules/Processor.bas" pos="219:17:19" line-data="    &#39;The File menu contains some abnormal operations (e.g. &quot;exit program&quot;) which require us to deal with their return">`e.g`</SwmToken>., <SwmToken path="Modules/Processor.bas" pos="1720:1:1" line-data="        FilterGridBlur">`FilterGridBlur`</SwmToken>) to keep old macros working or for internal use.

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

### Grid Blur Calculation and Application

```mermaid
%%{init: {"flowchart": {"defaultRenderer": "elk"}} }%%
flowchart TD
    node1["Prepare image data and variables"]
    click node1 openCode "Modules/Filters_Area.bas:277:312"
    node1 --> node2["Show 'Generating grids...' message"]
    click node2 openCode "Modules/Filters_Area.bas:279:279"
    node2 --> node3["Determine progress bar update frequency"]
    click node3 openCode "Modules/Filters_Area.bas:304:304"
    node3 --> loop1
    subgraph loop1["For each column"]
        loop1a["Calculate vertical color averages"]
        click loop1a openCode "Modules/Filters_Area.bas:314:327"
    end
    loop1 --> loop2
    subgraph loop2["For each row"]
        loop2a["Calculate horizontal color averages"]
        click loop2a openCode "Modules/Filters_Area.bas:330:343"
    end
    loop2 --> node4["Show 'Applying grid blur...' message"]
    click node4 openCode "Modules/Filters_Area.bas:345:345"
    node4 --> loop3
    subgraph loop3["For each pixel"]
        loop3a["Blend vertical and horizontal averages"]
        click loop3a openCode "Modules/Filters_Area.bas:348:355"
        loop3a --> loop3b{"Are color values > 255?"}
        click loop3b openCode "Modules/Filters_Area.bas:358:360"
        loop3b -->|"Yes"| loop3c["Clamp color values to 255"]
        click loop3c openCode "Modules/Filters_Area.bas:358:360"
        loop3b -->|"No"| loop3d["Assign blended color to pixel"]
        click loop3d openCode "Modules/Filters_Area.bas:363:365"
        loop3c --> loop3d
        loop3d --> loop3e{"User cancels?"}
        click loop3e openCode "Modules/Filters_Area.bas:369:369"
        loop3e -->|"Yes"| loop3f["Cancel operation"]
        click loop3f openCode "Modules/Filters_Area.bas:369:369"
        loop3e -->|"No"| loop3g["Update progress bar if needed"]
        click loop3g openCode "Modules/Filters_Area.bas:370:371"
    end
    loop3 --> node5["Deallocate image data"]
    click node5 openCode "Modules/Filters_Area.bas:375:375"
    node5 --> node6["Finalize and render image"]
    click node6 openCode "Modules/Filters_Area.bas:378:378"

classDef HeadingStyle fill:#777777,stroke:#333,stroke-width:2px;

%% Swimm:
%% %%{init: {"flowchart": {"defaultRenderer": "elk"}} }%%
%% flowchart TD
%%     node1["Prepare image data and variables"]
%%     click node1 openCode "<SwmPath>[Modules/Filters_Area.bas](Modules/Filters_Area.bas)</SwmPath>:277:312"
%%     node1 --> node2["Show 'Generating grids...' message"]
%%     click node2 openCode "<SwmPath>[Modules/Filters_Area.bas](Modules/Filters_Area.bas)</SwmPath>:279:279"
%%     node2 --> node3["Determine progress bar update frequency"]
%%     click node3 openCode "<SwmPath>[Modules/Filters_Area.bas](Modules/Filters_Area.bas)</SwmPath>:304:304"
%%     node3 --> loop1
%%     subgraph loop1["For each column"]
%%         loop1a["Calculate vertical color averages"]
%%         click loop1a openCode "<SwmPath>[Modules/Filters_Area.bas](Modules/Filters_Area.bas)</SwmPath>:314:327"
%%     end
%%     loop1 --> loop2
%%     subgraph loop2["For each row"]
%%         loop2a["Calculate horizontal color averages"]
%%         click loop2a openCode "<SwmPath>[Modules/Filters_Area.bas](Modules/Filters_Area.bas)</SwmPath>:330:343"
%%     end
%%     loop2 --> node4["Show 'Applying grid blur...' message"]
%%     click node4 openCode "<SwmPath>[Modules/Filters_Area.bas](Modules/Filters_Area.bas)</SwmPath>:345:345"
%%     node4 --> loop3
%%     subgraph loop3["For each pixel"]
%%         loop3a["Blend vertical and horizontal averages"]
%%         click loop3a openCode "<SwmPath>[Modules/Filters_Area.bas](Modules/Filters_Area.bas)</SwmPath>:348:355"
%%         loop3a --> loop3b{"Are color values > 255?"}
%%         click loop3b openCode "<SwmPath>[Modules/Filters_Area.bas](Modules/Filters_Area.bas)</SwmPath>:358:360"
%%         loop3b -->|"Yes"| loop3c["Clamp color values to 255"]
%%         click loop3c openCode "<SwmPath>[Modules/Filters_Area.bas](Modules/Filters_Area.bas)</SwmPath>:358:360"
%%         loop3b -->|"No"| loop3d["Assign blended color to pixel"]
%%         click loop3d openCode "<SwmPath>[Modules/Filters_Area.bas](Modules/Filters_Area.bas)</SwmPath>:363:365"
%%         loop3c --> loop3d
%%         loop3d --> loop3e{"User cancels?"}
%%         click loop3e openCode "<SwmPath>[Modules/Filters_Area.bas](Modules/Filters_Area.bas)</SwmPath>:369:369"
%%         loop3e -->|"Yes"| loop3f["Cancel operation"]
%%         click loop3f openCode "<SwmPath>[Modules/Filters_Area.bas](Modules/Filters_Area.bas)</SwmPath>:369:369"
%%         loop3e -->|"No"| loop3g["Update progress bar if needed"]
%%         click loop3g openCode "<SwmPath>[Modules/Filters_Area.bas](Modules/Filters_Area.bas)</SwmPath>:370:371"
%%     end
%%     loop3 --> node5["Deallocate image data"]
%%     click node5 openCode "<SwmPath>[Modules/Filters_Area.bas](Modules/Filters_Area.bas)</SwmPath>:375:375"
%%     node5 --> node6["Finalize and render image"]
%%     click node6 openCode "<SwmPath>[Modules/Filters_Area.bas](Modules/Filters_Area.bas)</SwmPath>:378:378"
%% 
%% classDef HeadingStyle fill:#777777,stroke:#333,stroke-width:2px;
```

<SwmSnippet path="/Modules/Filters_Area.bas" line="277">

---

In <SwmToken path="Modules/Filters_Area.bas" pos="277:4:4" line-data="Public Sub FilterGridBlur()">`FilterGridBlur`</SwmToken>, we prep the image data and calculate color sums for each vertical and horizontal line. These sums are used to average colors for the grid blur effect. Progress bar update logic is set up to keep things responsive, and we show a message to let the user know what's happening.

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

After prepping the color sums, <SwmToken path="Modules/Processor.bas" pos="1720:1:1" line-data="        FilterGridBlur">`FilterGridBlur`</SwmToken> loops over each pixel, averages the vertical and horizontal values, clamps them to 255, and writes the result back. Progress bar updates and ESC checks are done only when needed to keep things fast.

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

After <SwmPath>[Modules/ProgressBars.bas](Modules/ProgressBars.bas)</SwmPath>, we finalize the grid blur by committing the image changes and updating the UI so the user sees the result.

```visual basic
    'Safely deallocate imageData()
    workingDIB.UnwrapArrayFromDIB imageData
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering
    EffectPrep.FinalizeImageData

End Sub
```

---

</SwmSnippet>

### Dispatching Distort, Edge, Light, Noise, Pixelate, Render, Sharpen, Stylize, Transform, and Animation Effects

```mermaid
%%{init: {"flowchart": {"defaultRenderer": "elk"}} }%%
flowchart TD
  node1["User selects an effect or filter (processID)"] --> node2{"Is processID a recognized effect/filter?"}
  click node1 openCode "Modules/Processor.bas:1723:2014"
  node2 -->|"Yes"| node3{"Show dialog before applying effect? (raiseDialog)"}
  click node2 openCode "Modules/Processor.bas:1723:2014"
  node3 -->|"Yes"| node4["Show dialog for effect options"]
  click node3 openCode "Modules/Processor.bas:1725:2013"
  node3 -->|"No"| node5["Apply selected effect to image"]
  click node4 openCode "Modules/Processor.bas:1725:2013"
  click node5 openCode "Modules/Processor.bas:1725:2013"
  node4 --> node6["Return: Effect applied"]
  node5 --> node6
  click node6 openCode "Modules/Processor.bas:1726:2014"
  node2 -->|"No"| node7["No effect applied"]
  click node7 openCode "Modules/Processor.bas:2014:2025"
classDef HeadingStyle fill:#777777,stroke:#333,stroke-width:2px;

%% Swimm:
%% %%{init: {"flowchart": {"defaultRenderer": "elk"}} }%%
%% flowchart TD
%%   node1["User selects an effect or filter (<SwmToken path="Modules/Processor.bas" pos="89:8:8" line-data="Public Sub Process(ByVal processID As String, Optional raiseDialog As Boolean = False, Optional processParameters As String = vbNullString, Optional createUndo As PD_UndoType = UNDO_Nothing, Optional relevantTool As Long = -1, Optional recordAction As Boolean = True)">`processID`</SwmToken>)"] --> node2{"Is <SwmToken path="Modules/Processor.bas" pos="89:8:8" line-data="Public Sub Process(ByVal processID As String, Optional raiseDialog As Boolean = False, Optional processParameters As String = vbNullString, Optional createUndo As PD_UndoType = UNDO_Nothing, Optional relevantTool As Long = -1, Optional recordAction As Boolean = True)">`processID`</SwmToken> a recognized effect/filter?"}
%%   click node1 openCode "<SwmPath>[Modules/Processor.bas](Modules/Processor.bas)</SwmPath>:1723:2014"
%%   node2 -->|"Yes"| node3{"Show dialog before applying effect? (<SwmToken path="Modules/Processor.bas" pos="89:17:17" line-data="Public Sub Process(ByVal processID As String, Optional raiseDialog As Boolean = False, Optional processParameters As String = vbNullString, Optional createUndo As PD_UndoType = UNDO_Nothing, Optional relevantTool As Long = -1, Optional recordAction As Boolean = True)">`raiseDialog`</SwmToken>)"}
%%   click node2 openCode "<SwmPath>[Modules/Processor.bas](Modules/Processor.bas)</SwmPath>:1723:2014"
%%   node3 -->|"Yes"| node4["Show dialog for effect options"]
%%   click node3 openCode "<SwmPath>[Modules/Processor.bas](Modules/Processor.bas)</SwmPath>:1725:2013"
%%   node3 -->|"No"| node5["Apply selected effect to image"]
%%   click node4 openCode "<SwmPath>[Modules/Processor.bas](Modules/Processor.bas)</SwmPath>:1725:2013"
%%   click node5 openCode "<SwmPath>[Modules/Processor.bas](Modules/Processor.bas)</SwmPath>:1725:2013"
%%   node4 --> node6["Return: Effect applied"]
%%   node5 --> node6
%%   click node6 openCode "<SwmPath>[Modules/Processor.bas](Modules/Processor.bas)</SwmPath>:1726:2014"
%%   node2 -->|"No"| node7["No effect applied"]
%%   click node7 openCode "<SwmPath>[Modules/Processor.bas](Modules/Processor.bas)</SwmPath>:2014:2025"
%% classDef HeadingStyle fill:#777777,stroke:#333,stroke-width:2px;
```

<SwmSnippet path="/Modules/Processor.bas" line="1723">

---

After returning from <SwmPath>[Modules/Filters_Area.bas](Modules/Filters_Area.bas)</SwmPath>, <SwmToken path="Modules/Processor.bas" pos="1726:1:1" line-data="        Process_EffectsMenu = True">`Process_EffectsMenu`</SwmToken> runs through a big <SwmToken path="Modules/Processor.bas" pos="1724:1:1" line-data="    ElseIf Strings.StringsEqual(processID, &quot;Correct lens distortion&quot;, True) Then">`ElseIf`</SwmToken> chain to match the <SwmToken path="Modules/Processor.bas" pos="1724:7:7" line-data="    ElseIf Strings.StringsEqual(processID, &quot;Correct lens distortion&quot;, True) Then">`processID`</SwmToken> to the right effect. For each effect, if <SwmToken path="Modules/Processor.bas" pos="1725:3:3" line-data="        If raiseDialog Then ShowPDDialog vbModal, FormLensCorrect Else FormLensCorrect.CorrectLensDistortion processParameters">`raiseDialog`</SwmToken> is True, it calls <SwmToken path="Modules/Processor.bas" pos="1725:7:7" line-data="        If raiseDialog Then ShowPDDialog vbModal, FormLensCorrect Else FormLensCorrect.CorrectLensDistortion processParameters">`ShowPDDialog`</SwmToken> from <SwmPath>[Modules/Interface.bas](Modules/Interface.bas)</SwmPath> to pop up a UI for user input; otherwise, it runs the effect directly. This lets users interact with effects or run them automatically, depending on the context.

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

After returning from <SwmPath>[Modules/Interface.bas](Modules/Interface.bas)</SwmPath>, <SwmToken path="Modules/Processor.bas" pos="2021:1:1" line-data="        Process_EffectsMenu = True">`Process_EffectsMenu`</SwmToken> wraps up by handling Photoshop (<SwmToken path="Modules/Processor.bas" pos="2016:2:2" line-data="    &#39;8bf filters have a weird workflow because we simply call &quot;execute&quot; on the plugin but then all handling">`8bf`</SwmToken>) plugins as a special case, using a wrapper to delegate control. The dispatcher matches <SwmToken path="Modules/Processor.bas" pos="2019:7:7" line-data="    ElseIf Strings.StringsEqual(processID, &quot;Photoshop (8bf) plugin&quot;, True) Then">`processID`</SwmToken> to effects, shows dialogs if needed, or runs effects directly. Each branch sets the return value to True when handled.

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

## Tool Menu Command Dispatch and Macro Operations

```mermaid
%%{init: {"flowchart": {"defaultRenderer": "elk"}} }%%
flowchart TD
    node1["Receive image operation request"] --> node2{"Is operation found in tool menu?"}
    click node1 openCode "Modules/Processor.bas:258:259"
    node2 -->|"Yes"| node3["Execute tool menu operation"]
    click node2 openCode "Modules/Processor.bas:259:259"
    node3 --> node7["Update Undo/Redo and UI"]
    click node3 openCode "Modules/Processor.bas:259:259"
    node2 -->|"No"| node4{"Is processID a recognized legacy operation?"}
    click node4 openCode "Modules/Processor.bas:262:318"
    node4 -->|"Paint, Pencil, Clone, Gradient, Crop"| node5["Execute legacy paint/crop/gradient operation"]
    click node5 openCode "Modules/Processor.bas:269:296"
    node4 -->|"Fill tool"| node6{"Is macro or batch running?"}
    click node6 openCode "Modules/Processor.bas:285:288"
    node6 -->|"Yes"| node8["Apply fill from macro"]
    click node8 openCode "Modules/Processor.bas:286:287"
    node6 -->|"No"| node9["Skip fill macro"]
    click node9 openCode "Modules/Processor.bas:288:288"
    node4 -->|"Modify layer/text"| node10{"Is macro or batch running?"}
    click node10 openCode "Modules/Processor.bas:306:317"
    node10 -->|"Yes"| node11["Apply layer/text modification from macro"]
    click node11 openCode "Modules/Processor.bas:307:317"
    node10 -->|"No"| node12["Skip layer/text macro"]
    click node12 openCode "Modules/Processor.bas:317:317"
    node4 -->|"Do nothing"| node13["No operation performed"]
    click node13 openCode "Modules/Processor.bas:300:301"
    node4 -->|"Unknown"| node14["Notify user of unknown operation"]
    click node14 openCode "Modules/Processor.bas:325:325"
    node5 --> node7
    node8 --> node7
    node9 --> node7
    node11 --> node7
    node12 --> node7
    node13 --> node7
    node14 --> node7
    node7["Update Undo/Redo and UI"]
    click node7 openCode "Modules/Processor.bas:332:354"
    node7 --> node15{"Did operation change image?"}
    click node15 openCode "Modules/Processor.bas:343:344"
    node15 -->|"Yes"| node16["Notify UI of image change"]
    click node16 openCode "Modules/Processor.bas:344:344"
    node15 -->|"No"| node17["Skip UI notification"]
    click node17 openCode "Modules/Processor.bas:344:344"
    node16 --> node18{"Should timing report be shown?"}
    click node18 openCode "Modules/Processor.bas:337:338"
    node18 -->|"Yes"| node19["Show timing report"]
    click node19 openCode "Modules/Processor.bas:337:338"
    node18 -->|"No"| node20["Finish"]
    click node20 openCode "Modules/Processor.bas:354:354"
    node17 --> node20
classDef HeadingStyle fill:#777777,stroke:#333,stroke-width:2px;

%% Swimm:
%% %%{init: {"flowchart": {"defaultRenderer": "elk"}} }%%
%% flowchart TD
%%     node1["Receive image operation request"] --> node2{"Is operation found in tool menu?"}
%%     click node1 openCode "<SwmPath>[Modules/Processor.bas](Modules/Processor.bas)</SwmPath>:258:259"
%%     node2 -->|"Yes"| node3["Execute tool menu operation"]
%%     click node2 openCode "<SwmPath>[Modules/Processor.bas](Modules/Processor.bas)</SwmPath>:259:259"
%%     node3 --> node7["Update <SwmToken path="Modules/Processor.bas" pos="166:33:35" line-data="    &#39; if selections persist after a size change - and this is particularly relevant for the Undo/Redo engine,">`Undo/Redo`</SwmToken> and UI"]
%%     click node3 openCode "<SwmPath>[Modules/Processor.bas](Modules/Processor.bas)</SwmPath>:259:259"
%%     node2 -->|"No"| node4{"Is <SwmToken path="Modules/Processor.bas" pos="89:8:8" line-data="Public Sub Process(ByVal processID As String, Optional raiseDialog As Boolean = False, Optional processParameters As String = vbNullString, Optional createUndo As PD_UndoType = UNDO_Nothing, Optional relevantTool As Long = -1, Optional recordAction As Boolean = True)">`processID`</SwmToken> a recognized legacy operation?"}
%%     click node4 openCode "<SwmPath>[Modules/Processor.bas](Modules/Processor.bas)</SwmPath>:262:318"
%%     node4 -->|"Paint, Pencil, Clone, Gradient, Crop"| node5["Execute legacy paint/crop/gradient operation"]
%%     click node5 openCode "<SwmPath>[Modules/Processor.bas](Modules/Processor.bas)</SwmPath>:269:296"
%%     node4 -->|"Fill tool"| node6{"Is macro or batch running?"}
%%     click node6 openCode "<SwmPath>[Modules/Processor.bas](Modules/Processor.bas)</SwmPath>:285:288"
%%     node6 -->|"Yes"| node8["Apply fill from macro"]
%%     click node8 openCode "<SwmPath>[Modules/Processor.bas](Modules/Processor.bas)</SwmPath>:286:287"
%%     node6 -->|"No"| node9["Skip fill macro"]
%%     click node9 openCode "<SwmPath>[Modules/Processor.bas](Modules/Processor.bas)</SwmPath>:288:288"
%%     node4 -->|"Modify layer/text"| node10{"Is macro or batch running?"}
%%     click node10 openCode "<SwmPath>[Modules/Processor.bas](Modules/Processor.bas)</SwmPath>:306:317"
%%     node10 -->|"Yes"| node11["Apply layer/text modification from macro"]
%%     click node11 openCode "<SwmPath>[Modules/Processor.bas](Modules/Processor.bas)</SwmPath>:307:317"
%%     node10 -->|"No"| node12["Skip layer/text macro"]
%%     click node12 openCode "<SwmPath>[Modules/Processor.bas](Modules/Processor.bas)</SwmPath>:317:317"
%%     node4 -->|"Do nothing"| node13["No operation performed"]
%%     click node13 openCode "<SwmPath>[Modules/Processor.bas](Modules/Processor.bas)</SwmPath>:300:301"
%%     node4 -->|"Unknown"| node14["Notify user of unknown operation"]
%%     click node14 openCode "<SwmPath>[Modules/Processor.bas](Modules/Processor.bas)</SwmPath>:325:325"
%%     node5 --> node7
%%     node8 --> node7
%%     node9 --> node7
%%     node11 --> node7
%%     node12 --> node7
%%     node13 --> node7
%%     node14 --> node7
%%     node7["Update <SwmToken path="Modules/Processor.bas" pos="166:33:35" line-data="    &#39; if selections persist after a size change - and this is particularly relevant for the Undo/Redo engine,">`Undo/Redo`</SwmToken> and UI"]
%%     click node7 openCode "<SwmPath>[Modules/Processor.bas](Modules/Processor.bas)</SwmPath>:332:354"
%%     node7 --> node15{"Did operation change image?"}
%%     click node15 openCode "<SwmPath>[Modules/Processor.bas](Modules/Processor.bas)</SwmPath>:343:344"
%%     node15 -->|"Yes"| node16["Notify UI of image change"]
%%     click node16 openCode "<SwmPath>[Modules/Processor.bas](Modules/Processor.bas)</SwmPath>:344:344"
%%     node15 -->|"No"| node17["Skip UI notification"]
%%     click node17 openCode "<SwmPath>[Modules/Processor.bas](Modules/Processor.bas)</SwmPath>:344:344"
%%     node16 --> node18{"Should timing report be shown?"}
%%     click node18 openCode "<SwmPath>[Modules/Processor.bas](Modules/Processor.bas)</SwmPath>:337:338"
%%     node18 -->|"Yes"| node19["Show timing report"]
%%     click node19 openCode "<SwmPath>[Modules/Processor.bas](Modules/Processor.bas)</SwmPath>:337:338"
%%     node18 -->|"No"| node20["Finish"]
%%     click node20 openCode "<SwmPath>[Modules/Processor.bas](Modules/Processor.bas)</SwmPath>:354:354"
%%     node17 --> node20
%% classDef HeadingStyle fill:#777777,stroke:#333,stroke-width:2px;
```

<SwmSnippet path="/Modules/Processor.bas" line="258">

---

After returning from <SwmToken path="Modules/Processor.bas" pos="256:15:15" line-data="    If (Not processFound) Then processFound = Process_EffectsMenu(processID, raiseDialog, processParameters, createUndo, relevantTool, recordAction, returnDetails)">`Process_EffectsMenu`</SwmToken>, Process checks if the <SwmToken path="Modules/Processor.bas" pos="259:17:17" line-data="    If (Not processFound) Then processFound = Process_ToolsMenu(processID, raiseDialog, processParameters, createUndo, relevantTool, recordAction, returnDetails)">`processID`</SwmToken> matches any tool menu actions. If not already handled, it calls <SwmToken path="Modules/Processor.bas" pos="259:15:15" line-data="    If (Not processFound) Then processFound = Process_ToolsMenu(processID, raiseDialog, processParameters, createUndo, relevantTool, recordAction, returnDetails)">`Process_ToolsMenu`</SwmToken> next, which uses XML strings for flexible parameter handling and supports macro operations.

```visual basic
    'Tool menu operations have been successfully migrated to XML strings.  (None of their functions raise special return conditions, FYI.)
    If (Not processFound) Then processFound = Process_ToolsMenu(processID, raiseDialog, processParameters, createUndo, relevantTool, recordAction, returnDetails)
    
```

---

</SwmSnippet>

<SwmSnippet path="/Modules/Processor.bas" line="1619">

---

<SwmToken path="Modules/Processor.bas" pos="1619:4:4" line-data="Private Function Process_ToolsMenu(ByVal processID As String, Optional raiseDialog As Boolean = False, Optional processParameters As String = vbNullString, Optional createUndo As PD_UndoType = UNDO_Nothing, Optional relevantTool As Long = -1, Optional recordAction As Boolean = True, Optional ByRef returnDetails As String = vbNullString) As Boolean">`Process_ToolsMenu`</SwmToken> checks the <SwmToken path="Modules/Processor.bas" pos="1619:8:8" line-data="Private Function Process_ToolsMenu(ByVal processID As String, Optional raiseDialog As Boolean = False, Optional processParameters As String = vbNullString, Optional createUndo As PD_UndoType = UNDO_Nothing, Optional relevantTool As Long = -1, Optional recordAction As Boolean = True, Optional ByRef returnDetails As String = vbNullString) As Boolean">`processID`</SwmToken> for macro commands ('Start macro recording', 'Stop macro recording', 'Play macro') and calls the right Macros method. If matched, it marks the command as handled; otherwise, it does nothing.

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

After <SwmToken path="Modules/Processor.bas" pos="259:15:15" line-data="    If (Not processFound) Then processFound = Process_ToolsMenu(processID, raiseDialog, processParameters, createUndo, relevantTool, recordAction, returnDetails)">`Process_ToolsMenu`</SwmToken>, Process checks for legacy paint and fill operations, then looks for 'Modify layer' or 'Modify text layer' commands during macro playback. If found, it calls <SwmToken path="Modules/Processor.bas" pos="307:1:1" line-data="                MiniProcess_NDFX_MacroPlayback thisProcData, False &#39;Forward the command to a dedicated processor">`MiniProcess_NDFX_MacroPlayback`</SwmToken> to apply the recorded changes automatically.

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

<SwmToken path="Modules/Processor.bas" pos="664:4:4" line-data="Private Sub MiniProcess_NDFX_MacroPlayback(ByRef srcProcData As PD_ProcessCall, Optional ByVal useTextMode As Boolean = False)">`MiniProcess_NDFX_MacroPlayback`</SwmToken> updates layer properties based on macro parameters, handling text and generic properties as needed.

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

After <SwmToken path="Modules/Processor.bas" pos="307:1:1" line-data="                MiniProcess_NDFX_MacroPlayback thisProcData, False &#39;Forward the command to a dedicated processor">`MiniProcess_NDFX_MacroPlayback`</SwmToken>, Process acts as the central dispatcher for all actions. If it can't match the <SwmToken path="Modules/Processor.bas" pos="325:6:6" line-data="            If (LenB(processID) &lt;&gt; 0) Then PDMsgBox &quot;Unknown processor request submitted: %1&quot; &amp; vbCrLf &amp; vbCrLf &amp; &quot;Please report this bug via the Help -&gt; Submit Bug Report menu.&quot;, vbCritical Or vbOKOnly, &quot;Error&quot;, processID">`processID`</SwmToken> to any known handler, it calls <SwmPath>[Modules/Interface.bas](Modules/Interface.bas)</SwmPath> to show a message box and prompt the user to report the bug. This keeps the system robust and helps catch unhandled cases.

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

<SwmToken path="Modules/Interface.bas" pos="1618:4:4" line-data="Public Function PDMsgBox(ByVal pMessage As String, ByVal pButtons As VbMsgBoxStyle, ByVal pTitle As String, ParamArray ExtraText() As Variant) As VbMsgBoxResult">`PDMsgBox`</SwmToken> in <SwmPath>[Modules/Interface.bas](Modules/Interface.bas)</SwmPath> translates messages and titles if a language object is active, plugs in dynamic text, and shows a custom dialog. If the dialog fails, it falls back to the system message box. Cursor state is managed before and after the dialog to keep things smooth.

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

After <SwmPath>[Modules/Interface.bas](Modules/Interface.bas)</SwmPath>, Process checks if timing reports are enabled and if the action changed the image. If so, it calls <SwmToken path="Modules/Processor.bas" pos="337:17:17" line-data="    If g_DisplayTimingReports And (createUndo &lt;&gt; UNDO_Nothing) Then ReportProcessorTimeTaken m_ProcessingTime">`ReportProcessorTimeTaken`</SwmToken> to show how long the operation took. This gives users feedback on performance for real edits.

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

<SwmToken path="Modules/Processor.bas" pos="462:4:4" line-data="Public Sub ReportProcessorTimeTaken(ByVal srcStartTime As Currency)">`ReportProcessorTimeTaken`</SwmToken> calculates how long the action took, formats the result, translates the message, and shows it to the user using the Message function. This gives immediate feedback in the user's language.

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

After timing, Process finalizes undo/redo so edits can be reversed or repeated.

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

## <SwmToken path="Modules/Processor.bas" pos="166:33:35" line-data="    &#39; if selections persist after a size change - and this is particularly relevant for the Undo/Redo engine,">`Undo/Redo`</SwmToken> Finalization and Cancellation Handling

```mermaid
%%{init: {"flowchart": {"defaultRenderer": "elk"}} }%%
flowchart TD
    node1["Finalize processing action"] --> node2{"Did user cancel the action?"}
    click node1 openCode "Modules/Processor.bas:1290:1293"
    node2 -->|"Yes"| node3["Reset interface and allow future cancels"]
    click node2 openCode "Modules/Processor.bas:1293:1301"
    click node3 openCode "Modules/Processor.bas:1296:1300"
    node2 -->|"No"| node4{"Should record undo/redo?
(undo type ≠ 'Nothing', macro not playing, recording enabled)"}
    click node4 openCode "Modules/Processor.bas:1312:1338"
    node4 -->|"Yes"| node5{"Is an image active?"}
    click node5 openCode "Modules/Processor.bas:1313:1337"
    node5 -->|"Yes"| node6["Record undo/redo state
(Handle 'Fade' action if needed)"]
    click node6 openCode "Modules/Processor.bas:1318:1335"
    node5 -->|"No"| node9["Finish"]
    node4 -->|"No"| node9
    node6 --> node9["Finish"]
    click node9 openCode "Modules/Processor.bas:1342:1342"
classDef HeadingStyle fill:#777777,stroke:#333,stroke-width:2px;

%% Swimm:
%% %%{init: {"flowchart": {"defaultRenderer": "elk"}} }%%
%% flowchart TD
%%     node1["Finalize processing action"] --> node2{"Did user cancel the action?"}
%%     click node1 openCode "<SwmPath>[Modules/Processor.bas](Modules/Processor.bas)</SwmPath>:1290:1293"
%%     node2 -->|"Yes"| node3["Reset interface and allow future cancels"]
%%     click node2 openCode "<SwmPath>[Modules/Processor.bas](Modules/Processor.bas)</SwmPath>:1293:1301"
%%     click node3 openCode "<SwmPath>[Modules/Processor.bas](Modules/Processor.bas)</SwmPath>:1296:1300"
%%     node2 -->|"No"| node4{"Should record undo/redo?
%% (undo type ≠ 'Nothing', macro not playing, recording enabled)"}
%%     click node4 openCode "<SwmPath>[Modules/Processor.bas](Modules/Processor.bas)</SwmPath>:1312:1338"
%%     node4 -->|"Yes"| node5{"Is an image active?"}
%%     click node5 openCode "<SwmPath>[Modules/Processor.bas](Modules/Processor.bas)</SwmPath>:1313:1337"
%%     node5 -->|"Yes"| node6["Record undo/redo state
%% (Handle 'Fade' action if needed)"]
%%     click node6 openCode "<SwmPath>[Modules/Processor.bas](Modules/Processor.bas)</SwmPath>:1318:1335"
%%     node5 -->|"No"| node9["Finish"]
%%     node4 -->|"No"| node9
%%     node6 --> node9["Finish"]
%%     click node9 openCode "<SwmPath>[Modules/Processor.bas](Modules/Processor.bas)</SwmPath>:1342:1342"
%% classDef HeadingStyle fill:#777777,stroke:#333,stroke-width:2px;
```

<SwmSnippet path="/Modules/Processor.bas" line="1290">

---

In <SwmToken path="Modules/Processor.bas" pos="1290:4:4" line-data="Private Sub FinalizeUndoRedoState(ByRef srcProcData As PD_ProcessCall, ByRef targetImage As pdImage)">`FinalizeUndoRedoState`</SwmToken>, if the user cancels the action, we reset the progress bar UI, show a cancellation message via <SwmPath>[Modules/Interface.bas](Modules/Interface.bas)</SwmPath>, and reset the cancel flag. This keeps the UI clean and lets users cancel future actions.

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

After showing the cancel message in <SwmPath>[Modules/Interface.bas](Modules/Interface.bas)</SwmPath>, <SwmToken path="Modules/Processor.bas" pos="351:1:1" line-data="    FinalizeUndoRedoState thisProcData, PDImages.GetActiveImage">`FinalizeUndoRedoState`</SwmToken> checks if Undo data should be created. It figures out the affected layer (active layer or -1 for selection), handles special cases like 'Fade', and calls the <SwmToken path="Modules/Processor.bas" pos="1331:3:3" line-data="                    targetImage.UndoManager.FillDIBWithLastUndoCopy tmpDIB, affectedLayerID, , True">`UndoManager`</SwmToken> to save the Undo data.

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

## UI Finalization and Error Handling

```mermaid
%%{init: {"flowchart": {"defaultRenderer": "elk"}} }%%
flowchart TD
  node1["Finalize processing and mark processor as idle"]
  click node1 openCode "Modules/Processor.bas:356:359"
  node1 --> node2{"Undo action performed? (createUndo)"}
  click node2 openCode "Modules/Processor.bas:362:363"
  node2 -->|"Yes"| node3["Return focus to active image"]
  click node3 openCode "Modules/Processor.bas:363:363"
  node2 -->|"No"| node4{"Macro batch running? (Macros.GetMacroStatus)"}
  click node4 openCode "Modules/Processor.bas:370:383"
  node4 -->|"No"| node5{"Dialog shown? (raiseDialog)"}
  click node5 openCode "Modules/Processor.bas:373:381"
  node5 -->|"Yes"| node6{"Dialog canceled? (Interface.GetLastShowDialogResult = vbCancel)"}
  click node6 openCode "Modules/Processor.bas:376:376"
  node6 -->|"No"| node7["Sync UI to current image"]
  click node7 openCode "Modules/Processor.bas:376:376"
  node6 -->|"Yes"| node8["Skip UI sync"]
  click node8 openCode "Modules/Processor.bas:375:375"
  node5 -->|"No"| node7
  node4 -->|"Yes"| node8
  node3 --> node9["Re-enable main form and restore UI"]
  click node9 openCode "Modules/Processor.bas:389:389"
  node7 --> node9
  node8 --> node9
  node9 --> node10{"Update ready to install? (Updates.IsUpdateReadyToInstall)"}
  click node10 openCode "Modules/Processor.bas:392:392"
  node10 -->|"Yes"| node11["Display update notification"]
  click node11 openCode "Modules/Processor.bas:392:392"
  node10 -->|"No"| node12["Continue"]
  node11 --> node12
  node12 --> node13{"Did an error occur?"}
  click node13 openCode "Modules/Processor.bas:402:440"
  node13 -->|"Yes"| node14["Show error message box"]
  click node14 openCode "Modules/Processor.bas:443:443"
  node14 --> node15{"User agrees to file bug report? (msgReturn = vbYes)"}
  click node15 openCode "Modules/Processor.bas:446:446"
  node15 -->|"Yes"| node16["Guide user to bug report"]
  click node16 openCode "Modules/Processor.bas:1345:1353"
  node15 -->|"No"| node17["End"]
  node13 -->|"No"| node17
classDef HeadingStyle fill:#777777,stroke:#333,stroke-width:2px;

%% Swimm:
%% %%{init: {"flowchart": {"defaultRenderer": "elk"}} }%%
%% flowchart TD
%%   node1["Finalize processing and mark processor as idle"]
%%   click node1 openCode "<SwmPath>[Modules/Processor.bas](Modules/Processor.bas)</SwmPath>:356:359"
%%   node1 --> node2{"Undo action performed? (<SwmToken path="Modules/Processor.bas" pos="89:43:43" line-data="Public Sub Process(ByVal processID As String, Optional raiseDialog As Boolean = False, Optional processParameters As String = vbNullString, Optional createUndo As PD_UndoType = UNDO_Nothing, Optional relevantTool As Long = -1, Optional recordAction As Boolean = True)">`createUndo`</SwmToken>)"}
%%   click node2 openCode "<SwmPath>[Modules/Processor.bas](Modules/Processor.bas)</SwmPath>:362:363"
%%   node2 -->|"Yes"| node3["Return focus to active image"]
%%   click node3 openCode "<SwmPath>[Modules/Processor.bas](Modules/Processor.bas)</SwmPath>:363:363"
%%   node2 -->|"No"| node4{"Macro batch running? (<SwmToken path="Modules/Processor.bas" pos="285:5:7" line-data="            If ((Macros.GetMacroStatus = MacroPLAYBACK) Or (Macros.GetMacroStatus = MacroBATCH)) And (LenB(processParameters) &lt;&gt; 0) Then">`Macros.GetMacroStatus`</SwmToken>)"}
%%   click node4 openCode "<SwmPath>[Modules/Processor.bas](Modules/Processor.bas)</SwmPath>:370:383"
%%   node4 -->|"No"| node5{"Dialog shown? (<SwmToken path="Modules/Processor.bas" pos="89:17:17" line-data="Public Sub Process(ByVal processID As String, Optional raiseDialog As Boolean = False, Optional processParameters As String = vbNullString, Optional createUndo As PD_UndoType = UNDO_Nothing, Optional relevantTool As Long = -1, Optional recordAction As Boolean = True)">`raiseDialog`</SwmToken>)"}
%%   click node5 openCode "<SwmPath>[Modules/Processor.bas](Modules/Processor.bas)</SwmPath>:373:381"
%%   node5 -->|"Yes"| node6{"Dialog canceled? (<SwmToken path="Modules/Processor.bas" pos="376:6:8" line-data="                If Not (Interface.GetLastShowDialogResult = vbCancel) Then Interface.SyncInterfaceToCurrentImage">`Interface.GetLastShowDialogResult`</SwmToken> = <SwmToken path="Modules/Processor.bas" pos="376:12:12" line-data="                If Not (Interface.GetLastShowDialogResult = vbCancel) Then Interface.SyncInterfaceToCurrentImage">`vbCancel`</SwmToken>)"}
%%   click node6 openCode "<SwmPath>[Modules/Processor.bas](Modules/Processor.bas)</SwmPath>:376:376"
%%   node6 -->|"No"| node7["Sync UI to current image"]
%%   click node7 openCode "<SwmPath>[Modules/Processor.bas](Modules/Processor.bas)</SwmPath>:376:376"
%%   node6 -->|"Yes"| node8["Skip UI sync"]
%%   click node8 openCode "<SwmPath>[Modules/Processor.bas](Modules/Processor.bas)</SwmPath>:375:375"
%%   node5 -->|"No"| node7
%%   node4 -->|"Yes"| node8
%%   node3 --> node9["<SwmToken path="Modules/Processor.bas" pos="387:2:4" line-data="    &#39;Re-enable the main form and restore things like selection animations and proper control focus.">`Re-enable`</SwmToken> main form and restore UI"]
%%   click node9 openCode "<SwmPath>[Modules/Processor.bas](Modules/Processor.bas)</SwmPath>:389:389"
%%   node7 --> node9
%%   node8 --> node9
%%   node9 --> node10{"Update ready to install? (<SwmToken path="Modules/Processor.bas" pos="392:3:5" line-data="    If Updates.IsUpdateReadyToInstall() Then Updates.DisplayUpdateNotification">`Updates.IsUpdateReadyToInstall`</SwmToken>)"}
%%   click node10 openCode "<SwmPath>[Modules/Processor.bas](Modules/Processor.bas)</SwmPath>:392:392"
%%   node10 -->|"Yes"| node11["Display update notification"]
%%   click node11 openCode "<SwmPath>[Modules/Processor.bas](Modules/Processor.bas)</SwmPath>:392:392"
%%   node10 -->|"No"| node12["Continue"]
%%   node11 --> node12
%%   node12 --> node13{"Did an error occur?"}
%%   click node13 openCode "<SwmPath>[Modules/Processor.bas](Modules/Processor.bas)</SwmPath>:402:440"
%%   node13 -->|"Yes"| node14["Show error message box"]
%%   click node14 openCode "<SwmPath>[Modules/Processor.bas](Modules/Processor.bas)</SwmPath>:443:443"
%%   node14 --> node15{"User agrees to file bug report? (<SwmToken path="Modules/Processor.bas" pos="413:17:17" line-data="    Dim addInfo As String, mType As VbMsgBoxStyle, msgReturn As VbMsgBoxResult">`msgReturn`</SwmToken> = <SwmToken path="Modules/Processor.bas" pos="446:8:8" line-data="    If (msgReturn = vbYes) Then FileErrorReport Err.Number">`vbYes`</SwmToken>)"}
%%   click node15 openCode "<SwmPath>[Modules/Processor.bas](Modules/Processor.bas)</SwmPath>:446:446"
%%   node15 -->|"Yes"| node16["Guide user to bug report"]
%%   click node16 openCode "<SwmPath>[Modules/Processor.bas](Modules/Processor.bas)</SwmPath>:1345:1353"
%%   node15 -->|"No"| node17["End"]
%%   node13 -->|"No"| node17
%% classDef HeadingStyle fill:#777777,stroke:#333,stroke-width:2px;
```

<SwmSnippet path="/Modules/Processor.bas" line="356">

---

After <SwmToken path="Modules/Processor.bas" pos="351:1:1" line-data="    FinalizeUndoRedoState thisProcData, PDImages.GetActiveImage">`FinalizeUndoRedoState`</SwmToken>, Process marks the processor as idle and finalizes UI updates. If an <SwmToken path="Modules/Processor.bas" pos="365:28:30" line-data="    &#39;The interface will automatically be synched if an image is open and some undo-related action was applied (via the">`undo-related`</SwmToken> action was applied, it restores focus and syncs the interface. If not, it checks if a dialog was raised and syncs the UI as needed. Finally, it calls <SwmToken path="Modules/Processor.bas" pos="389:1:1" line-data="    SetProcessorUI_Idle processID, raiseDialog, processParameters, createUndo, relevantTool, recordAction">`SetProcessorUI_Idle`</SwmToken> to re-enable the main form and restore control focus.

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

After <SwmToken path="Modules/Processor.bas" pos="159:1:1" line-data="        SetProcessorUI_Idle processID, raiseDialog, processParameters, createUndo, relevantTool, recordAction">`SetProcessorUI_Idle`</SwmToken>, Process flushes any pending UI syncs by calling <SwmToken path="Modules/Processor.bas" pos="410:1:3" line-data="    Interface.SyncInterfaceToCurrentImage">`Interface.SyncInterfaceToCurrentImage`</SwmToken>. Then it handles error reporting, showing localized messages for out-of-memory or unknown errors using <SwmPath>[Modules/Interface.bas](Modules/Interface.bas)</SwmPath>.

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

Here, after returning from <SwmPath>[Modules/Interface.bas](Modules/Interface.bas)</SwmPath>, Process uses <SwmToken path="Modules/Processor.bas" pos="443:5:5" line-data="    msgReturn = PDMsgBox(&quot;PhotoDemon has experienced an error.  Details on the problem include:&quot; &amp; vbCrLf &amp; vbCrLf &amp; &quot;Error number %1&quot; &amp; vbCrLf &amp; &quot;Description: %2&quot; &amp; vbCrLf &amp; vbCrLf &amp; &quot;%3&quot;, mType, &quot;Error&quot;, Err.Number, Err.Description, addInfo)">`PDMsgBox`</SwmToken> to show a detailed error message to the user. This isn't just a plain message box—it pulls in translation, formatting, and UI consistency from <SwmPath>[Modules/Interface.bas](Modules/Interface.bas)</SwmPath>, so users get a clear, localized error report with all the relevant details (error number, description, extra info).

```visual basic
    'Create the message box to return the error information
    msgReturn = PDMsgBox("PhotoDemon has experienced an error.  Details on the problem include:" & vbCrLf & vbCrLf & "Error number %1" & vbCrLf & "Description: %2" & vbCrLf & vbCrLf & "%3", mType, "Error", Err.Number, Err.Description, addInfo)
    
```

---

</SwmSnippet>

<SwmSnippet path="/Modules/Processor.bas" line="445">

---

Next, after the error dialog from <SwmPath>[Modules/Interface.bas](Modules/Interface.bas)</SwmPath>, Process checks if the user agreed to file a bug report (clicked 'Yes'). If so, it calls <SwmToken path="Modules/Processor.bas" pos="446:13:13" line-data="    If (msgReturn = vbYes) Then FileErrorReport Err.Number">`FileErrorReport`</SwmToken> to walk them through reporting the issue. This keeps error reporting user-driven and avoids spamming the bug tracker with unwanted reports.

```visual basic
    'If the message box return value is "Yes", the user is willing to file a bug report.
    If (msgReturn = vbYes) Then FileErrorReport Err.Number
        
End Sub
```

---

</SwmSnippet>

<SwmSnippet path="/Modules/Processor.bas" line="1345">

---

<SwmToken path="Modules/Processor.bas" pos="1345:4:4" line-data="Private Sub FileErrorReport(ByVal errNumber As Long)">`FileErrorReport`</SwmToken> opens the <SwmToken path="Modules/Processor.bas" pos="1347:14:14" line-data="    &#39;Shell a browser window with the GitHub issue report form">`GitHub`</SwmToken> issues page in the user's browser so they can report the bug right away. Then it uses <SwmToken path="Modules/Processor.bas" pos="1351:1:1" line-data="    PDMsgBox &quot;PhotoDemon has automatically opened the bug report webpage for you.  Please click the &quot;&quot;New Issue&quot;&quot; button, then select &quot;&quot;Bug Report&quot;&quot;.  Answer the questions as best you can, and please include the following error number somewhere in your report: %1&quot; &amp; vbCrLf &amp; vbCrLf &amp; &quot;When finished, click the Submit New Issue button.  Thank you!&quot;, vbInformation Or vbOKOnly, &quot;Bug report instructions&quot;, errNumber">`PDMsgBox`</SwmToken> from <SwmPath>[Modules/Interface.bas](Modules/Interface.bas)</SwmPath> to show clear instructions—click 'New Issue', pick 'Bug Report', and include the error number—so the user knows exactly what to do next. This combo makes it easy for users to help improve the app.

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
