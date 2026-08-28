namespace WordGeminiFormula.AddIn.Ribbon
{
    internal static class RibbonXml
    {
        internal const string Value = @"<?xml version='1.0' encoding='UTF-8'?>
<customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui'>
  <ribbon>
    <tabs>
      <tab id='tabWordGeminiFormula' label='AI Formula'>
        <group id='grpOcr' label='Gemini OCR'>
          <button id='btnOcrImage'
                  label='Ảnh → Word'
                  size='large'
                  imageMso='PictureInsertFromFile'
                  screentip='OCR ảnh bằng Gemini'
                  supertip='Chuyển ảnh đề/tài liệu thành văn bản và công thức trong Word.'
                  onAction='OnOcrImage'/>
        </group>
        <group id='grpFormula' label='Công thức'>
          <button id='btnNormalizeAll'
                  label='Chuẩn hóa tất cả'
                  size='large'
                  imageMso='EquationInsertNew'
                  screentip='Chuẩn hóa công thức'
                  supertip='Chuyển các khối [[MATH]]...[[/MATH]] thành Word Equation Professional.'
                  onAction='OnNormalizeAll'/>
          <button id='btnNormalizeSelection'
                  label='Chuẩn hóa vùng chọn'
                  imageMso='EquationInsertNew'
                  onAction='OnNormalizeSelection'/>
        </group>
        <group id='grpSettings' label='Cấu hình'>
          <button id='btnSettings'
                  label='Settings'
                  size='large'
                  imageMso='FileProperties'
                  screentip='Cài đặt Gemini API'
                  supertip='Nhập API key và chọn Gemini model ngay trong Word.'
                  onAction='OnOpenSettings'/>
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>";
    }
}
