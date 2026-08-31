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
                  label='Ảnh → Word đẹp'
                  size='large'
                  imageMso='PictureInsertFromFile'
                  screentip='OCR ảnh bằng Gemini'
                  supertip='Nhận diện bố cục, văn bản, công thức, câu hỏi và giữ vùng hình/bảng khó OCR.'
                  onAction='OnOcrImage'/>
        </group>
        <group id='grpFormat' label='Format'>
          <button id='btnBeautify'
                  label='Làm đẹp format'
                  size='large'
                  imageMso='FormatPainter'
                  screentip='Làm đẹp tài liệu hiện tại'
                  supertip='Chuẩn hóa font, khoảng cách, tiêu đề, phần, câu hỏi, đáp án và footer của tài liệu Word.'
                  onAction='OnBeautifyFormat'/>
        </group>
        <group id='grpFormula' label='Công thức'>
          <button id='btnNormalizeAll'
                  label='Chuẩn hóa tất cả'
                  size='large'
                  imageMso='EquationInsertNew'
                  screentip='Chuẩn hóa công thức'
                  supertip='Chuyển các khối [[MATH]]...[[/MATH]] thành Word Equation Professional; vùng lỗi sẽ được tô vàng.'
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
                  screentip='Cài đặt Gemini API và format'
                  supertip='Nhập API key, chọn model, bật làm đẹp tự động và giữ vùng hình/bảng khó OCR.'
                  onAction='OnOpenSettings'/>
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>";
    }
}
