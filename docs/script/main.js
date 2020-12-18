const add_new_file = async file => {
  const resultList = document.querySelector('#resultList');
  const item = create_new_item(file);
  const tr = item.querySelector('tr');
  const nameEl = tr.querySelector('.result-name');
  const linkEl = tr.querySelector('.download-a');
  const dlBtn = tr.querySelector('.download-btn');
  resultList.appendChild(item);
  try {
    linkEl.href = await unlock(file);
    dlBtn.disabled = false;
  } catch(e) {
    nameEl.textContent += `(${e})`;
    dlBtn.querySelector('span.material-icons').textContent = 'report_problem';
  }
}

const create_new_item = file => {
  const template = document.querySelector('#productItem');
  const item = template.content.cloneNode(true);
  const tr = item.querySelector('tr');
  const nameEl = tr.querySelector('.result-name');
  const linkEl = tr.querySelector('.download-a');
  const closeBtn = tr.querySelector('.close-btn');
  nameEl.textContent = file.name;
  linkEl.download = file.name;
  closeBtn.addEventListener('click', e => {
    const item = e.currentTarget.closest('.result-item');
    const downloadAnker = tr.querySelector('.download-a');
    URL.revokeObjectURL(downloadAnker.href);
    item.remove();
  });
  return item;
}



const unlock = async file => {
  const rule = {
    'application/vnd.openxmlformats-officedocument.wordprocessingml.document': {
      'word/settings.xml': ['w:documentProtection']
    },
    'application/vnd.openxmlformats-officedocument.presentationml.presentation': {
      'ppt/presentation.xml': ['p:modifyVerifier']
    }
  };
  const target = rule[file.type];
  if (!target) throw 'FileType Error';
  const zip = await JSZip.loadAsync(file);
  for (const [path, tags] of Object.entries(target)) {
    const xmlSrc = await zip.file(path).async('string');
    const parser = new DOMParser();
    const doc = parser.parseFromString(xmlSrc, 'application/xml');
    for (const tag of tags)
      for (const elm of doc.getElementsByTagName(tag)) elm.remove();
    const serializer = new XMLSerializer();
    const xmlStr = serializer.serializeToString(doc);
    zip.file(path, xmlStr);
  }
  const blob = await zip.generateAsync({
    type: 'blob',
    mimeType: file.type
  });
  const url = URL.createObjectURL(blob);
  return url;
}


document.querySelector("#inputFiles").addEventListener('change', e => {
  const files = e.currentTarget.files;
  for (const file of files) {
    add_new_file(file);
  }
  e.currentTarget.value = '';
});