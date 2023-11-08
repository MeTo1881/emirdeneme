let fisListesi = [];

function tekilEkle() {
    let formAlani = document.getElementById('formAlani');
    formAlani.innerHTML = '';

    let form = document.createElement('form');
    form.id = 'formTekil';

    form.innerHTML = `
        <label for="tarih">Tarih:</label>
        <input type="date" id="tarih" name="tarih" required>

        <label for="fişNo">Fiş No:</label>
        <input type="text" id="fişNo" name="fişNo" required>

        <label for="firmaAdi">Firma Adı:</label>
        <input type="text" id="firmaAdi" name="firmaAdi" required>

        <label for="tutar">Tutar:</label>
        <input type="number" id="tutar" name="tutar" min="0" required>

        <label for="kdvOrani">KDV Oranı:</label>
        <select id="kdvOrani" name="kdvOrani">
            <option value="0.01">1%</option>
            <option value="0.10">10%</option>
            <option value="0.20">20%</option>
        </select>

        <button type="button" onclick="ekle()">Tekil Fiş Ekle</button>
    `;

    formAlani.appendChild(form);
}

function topluEkle() {
    let formAlani = document.getElementById('formAlani');
    formAlani.innerHTML = '';

    let topluFormSayisi = prompt('Kaç adet fiş eklemek istediğinizi girin:', '5');

    for (let i = 0; i < topluFormSayisi; i++) {
        let form = document.createElement('form');
        form.id = 'formToplu' + i;

        form.innerHTML = `
            <label for="tarih">Tarih:</label>
            <input type="date" id="tarih" name="tarih" required>

            <label for="fişNo">Fiş No:</label>
            <input type="text" id="fişNo" name="fişNo" required>

            <label for="firmaAdi">Firma Adı:</label>
            <input type="text" id="firmaAdi" name="firmaAdi" required>

            <label for="tutar">Tutar:</label>
            <input type="number" id="tutar" name="tutar" min="0" required>

            <label for="kdvOrani">KDV Oranı:</label>
            <select id="kdvOrani" name="kdvOrani">
                <option value="0.01">1%</option>
                <option value="0.10">10%</option>
                <option value="0.20">20%</option>
            </select>

            <button type="button" onclick="ekle()">Toplu Fiş Ekle</button>
        `;

        formAlani.appendChild(form);
    }
}

function ekle() {
    let formId = document.querySelector('form').id;

    if (formId.includes('Toplu')) {
        let forms = document.querySelectorAll('form');
        forms.forEach(form => {
            let tarih = form.querySelector('#tarih').value;
            let fişNo = form.querySelector('#fişNo').value;
            let firmaAdi = form.querySelector('#firmaAdi').value;
            let tutar = parseFloat(form.querySelector('#tutar').value);
            let kdvOrani = parseFloat(form.querySelector('#kdvOrani').value);

            let kdv = tutar * kdvOrani;
            let toplam = tutar + kdv;

            fisListesi.push({ tarih, fişNo, firmaAdi, tutar, kdv, toplam });
        });
    } else {
        let tarih = document.getElementById('tarih').value;
        let fişNo = document.getElementById('fişNo').value;
        let firmaAdi = document.getElementById('firmaAdi').value;
        let tutar = parseFloat(document.getElementById('tutar').value);
        let kdvOrani = parseFloat(document.getElementById('kdvOrani').value);

        let kdv = tutar * kdvOrani;
        let toplam = tutar + kdv;

        fisListesi.push({
            tarih,
            fişNo,
            firmaAdi,
            tutar,
            kdv,
            toplam
        });
    }

    fişListesiniGuncelle();

    document.getElementById('formAlani').innerHTML = '';
}

function fişListesiniGuncelle() {
    let fişListesiDiv = document.getElementById('fişListesi');
    fişListesiDiv.innerHTML = '';

    fisListesi.forEach((fiş, index) => {
        let fişDiv = document.createElement('div');
        fişDiv.classList.add('fis');
        fişDiv.innerHTML = `
            <p>Tarih: ${fiş.tarih}</p>
            <p>Fiş No: ${fiş.fişNo}</p>
            <p>Firma Adı: ${fiş.firmaAdi}</p>
            <p>Tutar: ${fiş.tutar}</p>
            <p>KDV: ${fiş.kdv}</p>
            <p>Toplam: ${fiş.toplam}</p>
        `;
               fişListesiDiv.appendChild(fişDiv);
    });
}

function excelOlustur() {
    let wb = XLSX.utils.book_new();
    wb.Props = {
        Title: 'Muhasebe Fiş Listesi',
        Subject: 'Muhasebe Fiş Listesi',
        Author: 'OpenAI',
        CreatedDate: new Date()
    };
    wb.SheetNames.push('Fiş Listesi');
    
    let ws_data = [];
    ws_data.push(["Tarih", "Fiş No", "Firma Adı", "Tutar", "KDV", "Toplam"]);

    fisListesi.forEach((fiş, index) => {
        ws_data.push([fiş.tarih, fiş.fişNo, fiş.firmaAdi, fiş.tutar, fiş.kdv, fiş.toplam]);
    });

    let ws = XLSX.utils.aoa_to_sheet(ws_data);
    wb.Sheets['Fiş Listesi'] = ws;

    let wbout = XLSX.write(wb, { bookType: 'xlsx', mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });

    function s2ab(s) {
        let buf = new ArrayBuffer(s.length);
        let view = new Uint8Array(buf);
        for (let i = 0; i < s.length; i++) view[i] = s.charCodeAt(i) & 0xFF;
        return buf;
    }

    saveAs(new Blob([s2ab(wbout)], { type: "application/octet-stream" }), 'Muhasebe_Fis_Listesi.xlsx');
}

