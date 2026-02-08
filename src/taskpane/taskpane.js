Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("run").onclick = runAI;
  }
});

async function runAI() {
  const prompt = document.getElementById("prompt").value;
  const loader = document.getElementById("loader");
  const resultArea = document.getElementById("result-area");
  const status = document.getElementById("status");

  if (!prompt) {
    alert("Iltimos, so'rov yozing!");
    return;
  }

  // Interfeysni tayyorlash
  loader.style.display = "block";
  resultArea.style.display = "none";
  status.innerText = "";

  try {
    // Bu yerda kelajakda API ulanadi
    // Hozircha simulyatsiya qilamiz
    await new Promise(resolve => setTimeout(resolve, 1500)); 

    status.innerText = "Nur IT AI Pro tahlili: Sizning so'rovingiz qabul qilindi. '" + 
                       prompt + "' mavzusida tahlil tayyorlanmoqda. " +
                       "\n\nMaslahat: Excel kataklaridagi ma'lumotlarni o'qish uchun diapazonni belgilang.";
    
    loader.style.display = "none";
    resultArea.style.display = "block";
  } catch (error) {
    status.innerText = "Xatolik yuz berdi: " + error.message;
    loader.style.display = "none";
    resultArea.style.display = "block";
  }
}