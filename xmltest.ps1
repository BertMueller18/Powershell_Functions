$xml = [xml]'
<Events>
<Event Definition="Validate" DLLPath="" DLLName="Helper.dll" DLLClass="HelpMain" DLLRoutine="pgFeatureInfoOnValidate_WriteToRegSelectedFeatures" InputParameters="pTreeViewFeatureTreeServerOS" RunOnce="no"/>
<Event Definition="Validate1" DLLPath="" DLLName="Helper.dll1" DLLClass="HelpMain1" DLLRoutine="pgFeatureInfoOnValidate_WriteToRegSelectedFeatures" InputParameters="pTreeViewFeatureTreeServerOS" RunOnce="no"/>
</Events>
'

Select-XML -xml $xml -xpath "//Event/@DllName"

$xml.SelectNodes('//Events/Event') | select DLLName