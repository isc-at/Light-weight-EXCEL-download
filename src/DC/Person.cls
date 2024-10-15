/// This sample persistent class represents a person.
Class DC.Person Extends (%Persistent, %Populate)
{

/// set sqlStatement="SELECT TOP 23 ID,Name,%ODBCOUT(DOB) DOB,SSN FROM Sample.Person"
/// Person's name.
Property Name As %String(POPSPEC = "Name()") [ Required ];

/// Person's Social Security number. This is validated using pattern match.
Property SSN As %String(PATTERN = "3N1""-""2N1""-""4N") [ Required ];

/// Person's Date of Birth.
Property DOB As %Date(POPSPEC = "Date()");

Storage Default
{
<Data name="PersonDefaultData">
<Value name="1">
<Value>%%CLASSNAME</Value>
</Value>
<Value name="2">
<Value>Name</Value>
</Value>
<Value name="3">
<Value>SSN</Value>
</Value>
<Value name="4">
<Value>DOB</Value>
</Value>
</Data>
<DataLocation>^DC.PersonD</DataLocation>
<DefaultData>PersonDefaultData</DefaultData>
<IdLocation>^DC.PersonD</IdLocation>
<IndexLocation>^DC.PersonI</IndexLocation>
<StreamLocation>^DC.PersonS</StreamLocation>
<Type>%Storage.Persistent</Type>
}

}