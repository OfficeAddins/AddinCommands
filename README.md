# AddinCommands
<B>I have created 3 simple manifest files and tried to include a built-in command using the 'OfficeControl' tag.

The command is displayed in PowerPoint, but not in Excel or Word.</B>


<A href="https://learn.microsoft.com/en-us/office/dev/add-ins/design/built-in-button-integration">learn.microsoft.com/en-us/office/dev/add-ins/design/built-in-button-integration</a>

The documentation says that this is only supported in PowerPoint at the moment.

<A href="https://learn.microsoft.com/en-us/javascript/api/manifest/group">learn.microsoft.com/en-us/javascript/api/manifest/group</a>

<img width="700" src="powerpoint-docs-screenshot.png">

```
<Group id="ExecuteButtonGroupId">
   <Label resid="ExecuteButton_Label"/>
   <Icon>
      <bt:Image size="16" resid="BET_16x16"/>
      <bt:Image size="32" resid="BET_32x32"/>
      <bt:Image size="80" resid="BET_80x80"/>
   </Icon>
   <Control xsi:type="Button" id="ExecuteButtonId">
      <Label resid="ExecuteButton_Label"/>
      <Supertip>
         <Title resid="ExecuteButton_Label"/>
         <Description resid="LS_Supertip"/>
      </Supertip>
      <Icon>
         <bt:Image size="16" resid="BET_16x16"/>
         <bt:Image size="32" resid="BET_32x32"/>
         <bt:Image size="80" resid="BET_80x80"/>
      </Icon>
      <Action xsi:type="ExecuteFunction">
         <FunctionName>ButtonPressedEvent</FunctionName>
      </Action>
   </Control>

   <OfficeControl id="Paste" />
</Group>
```


<H3>PowerPoint Desktop Application</H3>
When the manifest is side loaded into PowerPoint, the Paste command is displayed.
<img src="powerpoint-addincommands.png">

<H3>Excel Desktop Application</H3>
When the manifest is side loaded into Excel, the Paste command is NOT displayed.
<img src="excel-addincommands.png">

<H3>Word Desktop Application</H3>
When the manifest is side loaded into Word, the Paste command is NOT displayed.
<img src="word-addincommands.png">


<H3>Nice To Have - size="large"</H3>
Be able to display a built-in controls with a large size

```
<OfficeControl id="Paste" size="large" />
```

<H3>Nice To Have - on a menu</H3>
Be able to include built-in controls on a menu

```
<Control xsi:type="Menu">
  <Label/>
  <Icon/>
  <Items>
    <Item id="BUILTIN_ID">
      <OfficeControl id="Paste" />
    </Item>
  </Items>
```

<H3>Desktop Office Version</H3>
<img src="word-version.png">
