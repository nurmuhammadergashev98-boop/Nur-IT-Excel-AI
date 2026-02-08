Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("run").onclick = callGroqAI;
  }
});

async function callGroqAI() {
  const prompt = document.getElementById("prompt").value;
  const loader = document.getElementById("loader");
  const resultArea = document.getElementById("result-area");
  const status = document.getElementById("status");

  if (!prompt) {
    alert("Vazifa yozing!");
    return;
  }

  loader.style.display = "block";
  resultArea.style.display = "none";

  try {
    const response = await fetch("https://api.groq.com/openai/v1/chat/completions", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "Authorization": "Bearer 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"
      },
      body: JSON.stringify({
        model: "mixtral-8x7b-32768",
        messages: [
          { 
            role: "system", 
            content: "Sen Excel mutaxassisisan. Faqatgina so'ralgan formulani yoki natijani qaytar. Hech qanday ortiqcha gap yozma. Masalan: =SUM(A1:A10)" 
          },
          { role: "user", content: prompt }
        ]
      })
    });

    const data = await response.json();
    const aiFormula = data.choices[0].message.content.trim();

    // EXCELGA YOZISH QISMI
    await Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      // Agar AI formula qaytargan bo'lsa (boshida '=' bo'lsa)
      if (aiFormula.startsWith("=")) {
        range.formulas = [[aiFormula]];
      } else {
        range.values = [[aiFormula]];
      }
      await context.sync();
    });

    status.innerText = "Bajarildi! Katakka yozildi: " + aiFormula;
    loader.style.display = "none";
    resultArea.style.display = "block";

  } catch (error) {
    status.innerText = "Xatolik: " + error.message;
    loader.style.display = "none";
    resultArea.style.display = "block";
  }
}