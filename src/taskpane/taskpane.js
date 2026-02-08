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
    alert("Iltimos, savol yoki vazifa yozing!");
    return;
  }

  // Interfeysni yuklanish holatiga o'tkazish
  loader.style.display = "block";
  resultArea.style.display = "none";
  status.innerText = "";

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
            content: "Sen Nur IT kompaniyasi tomonidan yaratilgan Excel yordamchisisan. Foydalanuvchiga o'zbek tilida, aniq va professional javob ber." 
          },
          { role: "user", content: prompt }
        ],
        temperature: 0.7
      })
    });

    if (!response.ok) {
        throw new Error("API ulanishda xato: " + response.status);
    }

    const data = await response.json();
    const aiResponse = data.choices[0].message.content;

    // Natijani chiqarish
    status.innerText = aiResponse;
    loader.style.display = "none";
    resultArea.style.display = "block";

  } catch (error) {
    status.innerText = "Xatolik yuz berdi: " + error.message;
    loader.style.display = "none";
    resultArea.style.display = "block";
  }
}