(() => {
    const forms = document.querySelectorAll("form.export-form");
    if (!forms.length) return;  // stops errors if this file is loaded on another page

    const alertBox = document.getElementById("export-alert");
    const button = form.querySelector("button[type=submit]");

    function showAlert(type, message) {
        alertBox.className = "alert alert-" + type;
        alertBox.textContent = message;
        alertBox.hidden = false;
        setTimeout(() => { alertBox.hidden = true; }, 5000);
    }

    forms.addEventListener("submit", async (e) => {
        e.preventDefault();
        try {
            const res = await fetch(form.action, { method: "POST", body: new FormData(form) });
            if (!res.ok) throw new Error("Server returned " + res.status);

            const blob = await res.blob();
            const header = res.headers.get("Content-Disposition") || "";
            const match = header.match(/filename="?([^";]+)"?/);
            const filename = match ? match[1] : "report.pdf";

            const url = URL.createObjectURL(blob);
            const a = document.createElement("a");
            a.href = url;
            a.download = filename;
            document.body.appendChild(a);
            a.click();
            a.remove();
            setTimeout(() => URL.revokeObjectURL(url), 400);

            showAlert("success", "Report exported successfully: " + filename);
        } catch (err) {
            showAlert("danger", "Export failed. " + err.message);
        } finally {
            button.disabled = false;
            button.textContent = "Export Report";
        }
    });
})();