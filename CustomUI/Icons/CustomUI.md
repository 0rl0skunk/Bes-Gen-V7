# Custom UI
## CallBacks
Custom UI xml as reference for CallBack IDs:
``` xml
<customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui" onLoad="onLoad">
	<ribbon startFromScratch="false">
		<tabs>
			<tab id="customTab" label="Bes-Gen V7">
                <group id="customGroupPanels" label="Übersichten" getVisible="isVisibleGroup">
					<button id="Übersicht" label="Übersicht" image="bulletlist" size="large" onAction="onActionButton" getEnabled="isButtonEnabled"  supertip="Übersicht von Plänen im Projekt anzeigen und Pläne erstellen"/>
					<button id="Drucken" label="Publizieren" image="print" size="large" onAction="onActionButton" getEnabled="isButtonEnabled" supertip="Pläne plotten. [TinLine/AutoCAD Pläne]"/>
					<button id="Repair" label="Projekt Reparieren" image="repair" size="large" onAction="onActionButton" getEnabled="isButtonEnabled" supertip="TinLine Projekt bereinigen"/>
					<button id="Mail" label="E-Mail" image="microsoftoutlook" size="large" onAction="onActionButton" getEnabled="isButtonEnabled"  supertip="Pläne via E-Mail [Outlook] versenden. "/>
					<button id="Adresse" label="Adresse" image="addressbook" size="large" onAction="onActionButton" getEnabled="isButtonEnabled"  supertip="Adresse für Versand erfassen"/>
				</group>
				<group id="customGroupQuickAdd" label="Hinzufügen" getVisible="isVisibleGroup">
					<button id="Plan" label="Plan Hinzufügen" image="add" onAction="onActionButton" getEnabled="isButtonEnabled" supertip="Plan-Beschriftung Hinzufügen."/>
					<button id="Schema" label="Schema Hinzufügen" image="add" onAction="onActionButton" getEnabled="isButtonEnabled" supertip="Schema-Beschriftung Hinzufügen."/>
					<button id="Prinzip" label="Prinzip Hinzufügen" image="add" onAction="onActionButton" getEnabled="isButtonEnabled" supertip="Prinzip-Beschriftung Hinzufügen."/>
				</group>
                <group id="customGroupBuildings" label="Objektdaten" getVisible="isVisibleGroup">
					<button id="Objektdaten" label="Objektdaten erfassen" image="floorplan" size="large" onAction="onActionButton" getEnabled="isButtonEnabled" supertip="Objektdaten erfassen. [Gebäude, Geschoss]"/>
				</group>
				<group id="customGroupExplorer" getVisible="isVisibleGroup">
					<button id="CADFolder" label="CAD-Ordner" image="fileexplorernew" onAction="onActionButton" getEnabled="isButtonEnabled" supertip="Öffnet den Projektordner im Laufwerk H:\. [TinLine]"/>
					<button id="XREFFolder" label="XRef-Ordner" image="fileexplorernew" onAction="onActionButton" getEnabled="isButtonEnabled" supertip="Öffnet den XREF-Ordner vom TinLine im Laufwerk H:\. [TinLine]"/>
					<button id="SharePoint" label="SharePoint" image="sharepoint" onAction="onActionButton" getEnabled="isButtonEnabled" supertip="Öffnet den Projektordner im SharePoint. [Webansicht]"/>
				</group>
				<group id="customGroupHelp" label="Help" getVisible="isVisibleGroup">
					<button id="Version" label="Log" image="info" size="large" onAction="onActionButton" getEnabled="isButtonEnabled" supertip="Öffnet die Log-Datei welche abgelegt ist unter 'C:\Users\Public\Documents\TinLine\Bes-Gen_V7.log'"/>
					<button id="Chat" label="Support" image="onlinesupport" size="large" onAction="onActionButton" getEnabled="isButtonEnabled" visible="false" supertip="Öffnet den chat mit dem QS-Verantwortlichen"/>
					<button id="Bot" label="Bot" image="chatbot" size="large" onAction="onActionButton" getEnabled="isButtonEnabled" visible="false" supertip="Öffnet den ChatBot im Web."/>
					<button id="OneNote" label="OneNote" image="microsoftonenote" size="large" onAction="onActionButton" getEnabled="isButtonEnabled" supertip="Öffnet das OneNote mit Anleitungen etc."/>
					<button id="UserManual" label="Bedienungsanleitung" image="user-manual" size="large" onAction="onActionButton" getEnabled="isButtonEnabled" supertip="Öffnet die Interaktive Bedienungsanleitung"/>
				</group>
				<group id="customGroupCreateProject" label="Project" getVisible="isVisibleGroup">
					<button id="CADElektro" label="Elektro erstellen" image="projectsetup" size="large" onAction="onActionButton" getEnabled="isButtonEnabled" supertip="CAD-Projekt im Laufwerk H:\ erstellen. [TinLine]"/>
					<button id="Upgrade" label="Upgrade auf V7" image="windowsupdate" size="large" onAction="onActionButton" getEnabled="isButtonEnabled" supertip="Alte Versionen vom Beschriftungsgenerator auf Version 7 Updaten. !!!Die geöffnete Arbeitsmappe wird überschrieben!!!"/>
				</group>
				<group id="customGroupNoBesGen" label="Project" getVisible="isVisibleGroup">
					<button id="RedCross" image="error2" size="large"/>
					<labelControl id="LabelNoBesGen" label="Das geöffnete Dokument ist kein Beschriftungsgenerator." supertip="Update den Beschriftungsgenerator auf Version 7 oder frage beim QS-Verantworlichen um Hilfe."/>
				</group>
			</tab>
		</tabs>
	</ribbon>
</customUI>
```

```xml
<customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui" onLoad="onLoad">
	<ribbon startFromScratch="false">
        <tabs>
            <tab idMso="TabPictureToolsFormat" visible="false"/>
            <tab idMso="TabDrawingToolsFormat" visible="false"/>
            <tab idMso="TabInsert" visible="false"/>
            <tab idMso="TabPageLayoutExcel" visible="false"/>
            <tab idMso="TabView" visible="false"/>
            <tab idMso="TabFormulas" visible="false"/>
            <tab idMso="TabData" visible="false"/>
            <tab idMso="TabReview" visible="false"/>
            <tab idMso="TabSmartArtToolsDesign" visible="false"/>
            <tab idMso="TabSmartArtToolsFormat" visible="false"/>
            <tab idMso="TabChartToolsDesign" visible="false"/>
            <tab idMso="TabChartToolsLayout" visible="false"/>
            <tab idMso="TabChartToolsFormat" visible="false"/>
            <tab idMso="TabPivotTableToolsOptions" visible="false"/>
            <tab idMso="TabPivotTableToolsDesign" visible="false"/>
            <tab idMso="TabHome" visible="true"/>
            <tab idMso="TabHeaderAndFooterToolsDesign" visible="false"/>
            <tab idMso="TabAddIns" visible="false"/>
            <tab idMso="TabDeveloper" visible="false"/>
            <tab idMso="TabTableToolsDesignExcel" visible="false"/>
            <tab idMso="TabPrintPreview" visible="false"/>
            <tab idMso="TabInkToolsPens" visible="false"/>
            <tab idMso="TabPivotChartToolsAnalyze" visible="false"/>
            <tab idMso="TabPivotChartToolsDesign" visible="false"/>
            <tab idMso="TabPivotChartToolsLayout" visible="false"/>
            <tab idMso="TabPivotChartToolsFormat" visible="false"/>
            <tab idMso="TabAutomate" visible="false"/>
        </tabs>
    </ribbon>
</customUI>
```