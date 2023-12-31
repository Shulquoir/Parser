unit Parsing;

interface

uses
  SysUtils,  Classes;

type
  TParsing = class
  private
    const EOT_KEY_CODE = #4;
    const TAB_KEY_CODE = #9;
    const NUM_ONE = 1;
    const NUM_TWO = 2;
    const STR_ID = 'ID';
    const STR_ORIG = 'Orig';
    const STR_CURR = 'Curr';
    const STR_VALUES = '��������';
    const STR_TRANSLATION = '�������';
    const CHAR_COLON = ':';
    const CHAR_UPPER_A = #1040;
    const CHAR_LOWER_A = #1072;
    const CHAR_UPPER_YA = #1071;
    const CHAR_LOWER_YA = #1103;
    const CHAR_UPPER_YO = #1025;
    const CHAR_LOWER_YO = #1105;
  public
    class function ConvertDFNToTSV(const OpenedDFNFilePath: String; var OutputFileContent: TStringList): Boolean;
    class procedure ModifyDFN(const OpenedTSVFilePath: String; const OpenedDFNToModFilePath: String);
    class function ContainsCyrillicCharacters(const input: String): Boolean;
    class function IsCharCyrillic(c: Char): Boolean;
    class function ExtractField(const fields: TArray<String>; const fieldName: String): String;
  end;

implementation

/// <summary>
/// �������� ���� ����� DFN � ������ TSV �� ������ ���� � ��������� ��'��� TStringList.
/// </summary>
/// <param name="OpenedDFNFilePath">���� �� ��������� ����� DFN.</param>
/// <param name="OutputFileContent">��'��� TStringList, � ����� ���� ���������� ���� TSV �����.</param>
/// <returns>������� True, ���� ����������� ������� ������, ��� False � ������ �������.</returns>
class function TParsing.ConvertDFNToTSV(const OpenedDFNFilePath: String; var OutputFileContent: TStringList): Boolean;
var
  OriginalDFNFileContent: TStringList;

begin
  Result := false;

  try
    OriginalDFNFileContent := TStringList.Create; // ��������� ��'��� ���� TStringList
    OriginalDFNFileContent.LoadFromFile(OpenedDFNFilePath); // �������� ������� � �����
    OutputFileContent.Add(STR_ID + TAB_KEY_CODE + STR_VALUES + TAB_KEY_CODE + STR_TRANSLATION); // ������ ������ ����� � ��������� ���� TSV

    for var i := 0 to OriginalDFNFileContent.Count - NUM_ONE do // ���������� ����� � ����
    begin
      var line := OriginalDFNFileContent[i]; // ��������� ����� ��� ���� ������ ��� �����

      if ContainsCyrillicCharacters(line) and  // ���������� �� ��������� �������� ������� �� ��������� ��������
      (Pos(STR_ID, line) > 0) and (Pos(STR_ORIG, line) > 0) and (Pos(STR_CURR, line) > 0) then // �����
      begin     // ���������� �������� �� ������� � ����� ������ ����� ���� 'ID', 'Orig', 'Curr'
        var fields := line.Split([EOT_KEY_CODE]); // ��������� ����� �� ����, ���� ��������� �������� EOT
          var id := ExtractField(fields, STR_ID); // �� ��������� �������� �������, � ��� �������� ���� ��
          var orig := ExtractField(fields, STR_ORIG); // ������������� 'ID'/'Orig'/'Curr', �������� ����
          var curr := ExtractField(fields, STR_CURR); // ���������� ���� ��� ��������� �����, ���� ��� ��������
        OutputFileContent.Add(id + TAB_KEY_CODE + orig + TAB_KEY_CODE + curr);
      end;
    end;

    Result := true;

  finally
    FreeAndNil(OriginalDFNFileContent); // ��������� ���� ��ﳿ ��������� �����
  end;
end;

//------------------------------------------------------------------------------

/// <summary>
/// �������� ���� DFN ����� �� ����� ��������� TSV ����� �� ������ ���� � ��������� DFN ����.
/// </summary>
/// <param name="OpenedTSVFilePath">���� �� ��������� TSV �����.</param>
/// <param name="OpenedDFNToModFilePath">���� �� ��������� DFN ����� ��� �����������.</param>
class procedure TParsing.ModifyDFN(const OpenedTSVFilePath: String; const OpenedDFNToModFilePath: String);
var
  OriginalTSVFileContent, OriginalDFNToModFileContent: TStringList;
  tsvIDField, dfnIDField: String;

begin
  try
    OriginalTSVFileContent := TStringList.Create;
    OriginalDFNToModFileContent := TStringList.Create;

    OriginalTSVFileContent.LoadFromFile(OpenedTSVFilePath); // �������� ������� � �����
    OriginalDFNToModFileContent.LoadFromFile(OpenedDFNToModFilePath); // �������� ������� � �����

    for var TSVLine := 0 to OriginalTSVFileContent.Count - NUM_ONE do // ���������� ����� � ���� TSV
    begin
      tsvIDField := OriginalTSVFileContent[TSVLine].Split([TAB_KEY_CODE])[0]; // �������� � ����� ����� ���� �
                                                                                              // ����� TSV �����
      for var DFNLine := 0 to OriginalDFNToModFileContent.Count - NUM_ONE do // ���������� ����� � ���� DFN
      begin
        dfnIDField := ExtractField(OriginalDFNToModFileContent[DFNLine].Split([EOT_KEY_CODE]), STR_ID); //��������
                                                                                             // ���� "ID" DFN �����
        if Pos(tsvIDField, dfnIDField) = NUM_ONE then // ��������, �� ���� ID � DFN ���� ��������� ������� ���� TSV �����
        begin // ���� ���� �������� ���������, �� ������ ���� "Curr" � ����� DFN �����
          var OriginalDFNLine := OriginalDFNToModFileContent[DFNLine]; // ������� ����� DFN ����� � ����������
          var DFNFields := OriginalDFNToModFileContent[DFNLine].Split([EOT_KEY_CODE]); // ���� ���� ����� �
                                                                                       // ���������� ����� ID
          for var CurrField := 0 to Length(DFNFields) - NUM_ONE do // ���������� ���� ����� � ����� ���� ��������
          begin      // ��������� ���� "ID" ��� ����������� ������ ���� � ����� �������� ����� ���� "Curr",
                                            // ����������� ���� ��� ����� ��� ���������� ���������� ����
            if Pos(STR_CURR, DFNFields[CurrField]) >= NUM_ONE then // ������������ �������� ���� �� ���� 'Curr:'
            begin
              var DFNCurr := OriginalDFNToModFileContent[DFNLine].Split([EOT_KEY_CODE])[CurrField]; // �������� ���� Curr DFN �����
              var TSVCurr := OriginalTSVFileContent[TSVLine].Split([TAB_KEY_CODE])[NUM_TWO]; // �������� ���� Curr TSV �����
              OriginalDFNLine := OriginalDFNLine.Replace(DFNCurr, STR_CURR + CHAR_COLON + QuotedStr(TSVCurr)); // � ����� � ����������
              // ����� "ID" �������� ���� ���� "Curr", ���������� ����� ����� ������� CurrField, �� 3 ���� � ����� TSV
              OriginalDFNToModFileContent[DFNLine] := OriginalDFNLine; // ����� ������������ ����� ������������ ��ﳺ� �����
              Break; // ����� � �����, ���� �������� �������� ����
            end;
          end;

          Break; // ����� � �����, ���� �������� ��������� ���� ID, ���������� �� ���������� ����� � ���� TSV

        end;
      end;
    end;

    OriginalDFNToModFileContent.SaveToFile(OpenedDFNToModFilePath); // �������� ���� � ���� DFN

  finally
    FreeAndNil(OriginalTSVFileContent);
    FreeAndNil(OriginalDFNToModFileContent);
  end;
end;

//------------------------------------------------------------------------------

/// <summary>
/// ��������, �� ������ ����� ��������� �������.
/// </summary>
/// <param name="input">����� ��� ��������.</param>
/// <returns>True, ���� ����� ������ ��������� �������, ������ - False.</returns>
class function TParsing.ContainsCyrillicCharacters(const input: String): Boolean;
begin
  Result := False; // �������� � ��������� �������� false, ���� �� ���� �������� ��������
  for var i := NUM_ONE to Length(input) do // ���������� ����� � ���������� �����
  begin
    if IsCharCyrillic(input[i]) then // ������������� �� ���� ������� ��� �������� �� � ����� ���������
    begin
      Result := True; // ���� ���� �������� ����������� ������� ��������, ����������� � ��������� IsCharCyrillic
      Exit; // ��������� ��������� ������� � ����������� true
    end;
  end;
end;

//------------------------------------------------------------------------------

/// <summary>
/// ��������, �� � ������ ����������.
/// </summary>
/// <param name="c">������ ��� ��������.</param>
/// <returns>True, ���� ������ � ����������, ������ - False.</returns>
class function TParsing.IsCharCyrillic(c: Char): Boolean;
begin
  Result := (c >= CHAR_UPPER_A) and (c <= CHAR_UPPER_YA) or (c >= CHAR_LOWER_A) and
  (c <= CHAR_LOWER_YA) or (c = CHAR_UPPER_YO) or (c = CHAR_LOWER_YO);
end;

//------------------------------------------------------------------------------

/// <summary>
/// ��������� ���� ���� �� ���� ������.
/// </summary>
/// <param name="fields">����� ����, ����� ���� ����������� �����.</param>
/// <param name="fieldName">����� ����, ��� ��������� ������.</param>
/// <returns>���� ����, ���� ��������, ��� �������� �����, ���� �� ��������.</returns>
class function TParsing.ExtractField(const fields: TArray<String>; const fieldName: String): String;
var
  fieldValue: String;

begin
  for var field in fields do // ���������� ���� � �����
  begin
    if Pos(fieldName, field) > 0 then // ��������, �� �������� � ��� ������ ����� ����
    begin
      fieldValue := (Copy(field, Pos(fieldName, field) + Length(fieldName) + NUM_ONE, MaxInt));  // ���� ��������,
                              // �� ������� ���� ���� - ��� ����� ����, ������� ':' �� ���� ���� �� ���� ����
      if (Length(fieldValue) >= NUM_TWO) and (fieldValue[NUM_ONE] = '''') and  // �������� �� ������ ���� ����
      (fieldValue[Length(fieldValue)] = '''') then                        // � ���� �� �� ������� �������� �����
        Result := Copy(fieldValue, NUM_TWO, Length(fieldValue) - NUM_TWO) // ���� ������, �� �������
      else
        Result := fieldValue; // ���� ��, �� ������ �� ���������
      Break; // �������� � ����� ��� ����������� �������� ����� ����
    end
    else
      Result := ''; // ���� �� ���� �������� ��'� ����
  end;
end;

end.
