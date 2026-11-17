const fs = require('fs');

let xml = fs.readFileSync('G:\\1Click Build\\Source\\customUI.xml', 'utf8');

// 1. Add Tác giả to Home Tab if not present
if (!xml.includes('id="btnTacGiaHomeWord"')) {
  const homeTarget = '<button id="btnCapNhatHome"';
  const homeBtn = `<button id="btnTacGiaHomeWord" label="Tác giả" imageMso="ContactCard" screentip="Thông tin tác giả &amp; Bản quyền" onAction="RibbonThongTinTacGia" />\n          `;
  if (xml.includes(homeTarget)) {
    xml = xml.replace(homeTarget, homeBtn + homeTarget);
    console.log('Added btnTacGiaHomeWord to TabHome');
  }
}

// 2. Add grpTacGiaWord to tabTheThuc
if (!xml.includes('id="grpTacGiaWord"')) {
  const authorGroup = `
        <group id="grpTacGiaWord" label="Tác Giả &amp; Hỗ Trợ">
          <button id="btnThongTinTacGiaWord"
                  label="Thông Tin Tác Giả"
                  size="large"
                  imageMso="ContactCard"
                  screentip="Thông tin tác giả &amp; Bản quyền"
                  supertip="1Click Office 3.0 Enterprise&#10;Tác giả: Trần Quốc Được&#10;Đơn vị: Trung tâm Dịch vụ Tổng hợp Phường Tân Khai&#10;Bản quyền: Vĩnh viễn (Đã kích hoạt)."
                  onAction="RibbonThongTinTacGia" />
          <button id="btnKiemTraCapNhatWord"
                  label="Kiểm tra cập nhật"
                  imageMso="SyncStatusNormal"
                  screentip="Kiểm tra phiên bản mới từ máy chủ GitHub"
                  supertip="Tự động kiểm tra phiên bản mới từ GitHub và cập nhật nếu có."
                  onAction="RibbonKiemTraCapNhat" />
          <button id="btnHuongDanWord"
                  label="Hướng dẫn sử dụng"
                  imageMso="Help"
                  screentip="Xem tài liệu hướng dẫn sử dụng 1Click Office"
                  onAction="RibbonHuongDanSuDung" />
        </group>
      </tab>
    </tabs>`;

  // Replace '</tab>\n    </tabs>' with authorGroup
  const regex = /<\/tab>\s*<\/tabs>/;
  if (regex.test(xml)) {
    xml = xml.replace(regex, authorGroup);
    console.log('Successfully inserted grpTacGiaWord before </tabs>');
  } else {
    console.error('Could not find </tab></tabs> in customUI.xml!');
  }
}

fs.writeFileSync('G:\\1Click Build\\Source\\customUI.xml', xml, 'utf8');
console.log('Saved G:\\1Click Build\\Source\\customUI.xml');
