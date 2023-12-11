# Custom UI
## CallBacks
Custom UI xml as reference for CallBack IDs:
``` xml
<customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui">
	<ribbon startFromScratch="false">
		<tabs>
			<tab id="customTab" label="Bes-Gen V7">
                <group id="customGroupPanels" label="Panels">
					<button id="Drucken" label="Publizieren" image="print" size="large" onAction="CallbackUI" />
					<button id="Repair" label="Projekt Reparieren" image="repair" size="large" onAction="CallbackUI" />
					<button id="Übersicht" label="Übersicht" image="bulletlist" size="large" onAction="CallbackUI" />
				</group>
				<group id="customGroupSIA" label="Panels">
				    <button id="LockProjekt" label="Entsperren" image="lockopen" size="large" onAction="CallbackUI" />
				    <editBox id="Projektnummer" label="Projektnummer" onChange="CallbackUIText" getText="CallBackGetText"/>
				    <editBox id="Projektname" label="Projektname" onChange="CallbackUIText" getText="CallBackGetText"/>
					<comboBox id="comboBoxProjektphase" label="Projektphase" onChange="CallbackUIText" getText="CallBackGetText">
                        <item id="SIA31" label="Vorprojekt" />
                        <item id="SIA32" label="Bauprojekt" />
                        <item id="SIA41" label="Ausschreibung / SUBMISSION" />
                        <item id="SIA52" label="Ausführung" />
                        <item id="SIA53" label="Abschluss / REVISION" />
                    </comboBox>
                </group> 
                <group id="customGroupBuildings" label="Objektdaten">
					<button id="Objektdaten" label="Objektdaten erfassen" image="floorplan" size="large" onAction="CallbackUI" />
				</group>
				<group id="customGroupExplorer" label="Objektdaten">
					<button id="CADFolder" label="CAD-Ordner öffnen" image="fileexplorernew" size="large" onAction="CallbackUI" />
					<button id="SharePoint" label="SharePoint öffnen" image="sharepoint" size="large" onAction="CallbackUI" />
				</group>
				<group id="customGroupHelp" label="Help">
					<button id="Version" label="Versions Info" image="info" size="large" onAction="CallbackUI" />
					<button id="Chat" label="Support" image="onlineSupport" size="large" onAction="CallbackUI" />
					<button id="Bot" label="Bot" image="chatbot" size="large" onAction="CallbackUI" />
				</group>
			</tab>
		</tabs>
	</ribbon>
</customUI>
```