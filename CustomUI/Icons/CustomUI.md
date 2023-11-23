# Custom UI
## CallBacks
Custom UI xml as reference for CallBack IDs:
``` xml
<customUI xmlns="http://schemas.microsoft.com/office/2006/01/customui">
	<ribbon startFromScratch="false">
		<tabs>
			<tab id="customTab" label="Bes-Gen V7">
                <group id="customGroupPanels" label="Panels">
					<button id="Drucken" label="Publizieren" image="print" size="large" onAction="CallbackUI" />
					<button id="Repair" label="Projekt Reparieren" image="repair" size="large" onAction="CallbackUI" />
					<button id="Übersicht" label="Übersicht" image="bulletList" size="large" onAction="CallbackUI" />
				</group>
				<group id="customGroupSIA" label="Panels">
				    <editBox id="Projektname" label="Projektname" onAction="CallbackUI"/>
				    <editBox id="Projektbezeichnung" label="Projektbezeichnung" onAction="CallbackUI"/>
					<comboBox id="comboBoxProjektphase" label="Projektphase" onAction="CallbackUI">
                        <item id="SIA52" label="Ausführung" />
                        <item id="SIA31" label="Bauprojekt" />
                        <item id="SIA42" label="Ausschreibung" />
                    </comboBox>
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