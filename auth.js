(function () {
  const decode = value => Uint8Array.from(atob(value), character => character.charCodeAt(0));
  const derive = async (password, salt, iterations, usage) => {
    const material = await crypto.subtle.importKey("raw", new TextEncoder().encode(password), "PBKDF2", false, [usage === "bits" ? "deriveBits" : "deriveKey"]);
    const params = {name: "PBKDF2", hash: "SHA-256", salt, iterations};
    return usage === "bits" ? new Uint8Array(await crypto.subtle.deriveBits(params, material, 256)) : crypto.subtle.deriveKey(params, material, {name: "AES-GCM", length: 256}, false, ["decrypt"]);
  };
  const sameBytes = (left, right) => left.length === right.length && left.reduce((different, byte, index) => different | (byte ^ right[index]), 0) === 0;
  const style = document.createElement("style");
  style.textContent = `.auth-lock{position:fixed;inset:0;z-index:1000;display:grid;place-items:center;padding:20px;background:#f4f7f5}.auth-box{width:min(480px,100%);padding:28px;background:#fff;border:1px solid #dce4df;border-radius:18px;box-shadow:0 24px 70px rgba(0,0,0,.18);color:#15231d;font:15px/1.45 Inter,Segoe UI,Arial,sans-serif}.auth-box h1{margin:0 0 8px;font-size:28px}.auth-box p{color:#65736c}.auth-box label{display:block;font-weight:700;margin:18px 0 7px}.auth-box input[type=password]{width:100%;padding:12px;border:1px solid #cbd6d0;border-radius:10px;font:inherit}.auth-box button{width:100%;margin-top:12px;padding:11px;border:0;border-radius:10px;background:#15231d;color:#fff;font:700 15px inherit;cursor:pointer}.auth-box button:disabled{opacity:.55}.auth-error{min-height:22px;margin-top:10px;color:#b83a3a!important}.data-step[hidden]{display:none}.drop-zone{display:grid;place-items:center;min-height:170px;margin-top:18px;padding:22px;border:2px dashed #9bb4a8;border-radius:14px;background:#f7faf8;text-align:center;cursor:pointer}.drop-zone.drag{border-color:#087a55;background:#dff3ea}.drop-zone b{display:block;font-size:18px}.drop-zone input{position:absolute;width:1px;height:1px;opacity:0}`;
  document.head.appendChild(style);

  async function verify(password) {
    const config = globalThis.KSK_AUTH_CONFIG;
    if (!config) throw new Error("Thiếu cấu hình xác thực.");
    const actual = await derive(password, decode(config.salt), config.iterations, "bits");
    return sameBytes(actual, decode(config.hash));
  }
  async function decryptFile(file, password) {
    const payload = JSON.parse(await file.text());
    if (payload.version !== 1 || !payload.salt || !payload.iv || !payload.ciphertext) throw new Error("File dữ liệu không đúng định dạng.");
    const key = await derive(password, decode(payload.salt), payload.iterations, "key");
    const plaintext = await crypto.subtle.decrypt({name: "AES-GCM", iv: decode(payload.iv)}, key, decode(payload.ciphertext));
    return JSON.parse(new TextDecoder().decode(plaintext));
  }

  function unlock() {
    return new Promise(resolve => {
      let password = "";
      const lock = document.createElement("div");
      lock.className = "auth-lock";
      lock.innerHTML = `<section class="auth-box" role="dialog" aria-modal="true" aria-labelledby="authTitle"><div class="password-step"><form><h1 id="authTitle">Dữ liệu KSK</h1><p>Nhập mật khẩu truy cập. Trang web chỉ lưu giá trị băm, không chứa dữ liệu hồ sơ.</p><label for="accessPassword">Mật khẩu</label><input id="accessPassword" type="password" autocomplete="current-password" required autofocus><button type="submit">Đăng nhập</button><p class="auth-error" role="alert" aria-live="polite"></p></form></div><div class="data-step" hidden><h1>Nạp dữ liệu mã hóa</h1><p>Đăng nhập thành công. Kéo thả file <b>.ksk</b> được cấp riêng để mở dashboard.</p><label class="drop-zone" for="encryptedData"><span><b>Thả file dữ liệu vào đây</b>hoặc nhấn để chọn file</span><input id="encryptedData" type="file" accept=".ksk,.enc,application/json"></label><p class="auth-error file-error" role="alert" aria-live="polite"></p></div></section>`;
      document.body.appendChild(lock);
      const passwordStep = lock.querySelector(".password-step"), dataStep = lock.querySelector(".data-step"), form = lock.querySelector("form"), input = lock.querySelector("#accessPassword"), button = form.querySelector("button"), error = form.querySelector(".auth-error"), fileInput = lock.querySelector("#encryptedData"), dropZone = lock.querySelector(".drop-zone"), fileError = lock.querySelector(".file-error");
      form.onsubmit = async event => {
        event.preventDefault(); button.disabled = true; error.textContent = "Đang kiểm tra…";
        try {
          if (!await verify(input.value)) throw new Error("Mật khẩu không đúng.");
          password = input.value; input.value = ""; passwordStep.hidden = true; dataStep.hidden = false; fileInput.focus();
        } catch (cause) {
          input.value = ""; input.focus(); button.disabled = false; error.textContent = cause.message;
        }
      };
      const load = async file => {
        if (!file) return;
        fileError.textContent = `Đang giải mã ${file.name}…`;
        try {
          const data = await decryptFile(file, password);
          password = ""; fileInput.value = ""; lock.remove(); style.remove(); resolve(data);
        } catch (cause) {
          console.error(cause); fileInput.value = "";
          fileError.textContent = cause instanceof DOMException ? "Không giải mã được: sai mật khẩu hoặc file dữ liệu không hợp lệ." : cause.message;
        }
      };
      fileInput.onchange = () => load(fileInput.files[0]);
      for (const eventName of ["dragenter", "dragover"]) dropZone.addEventListener(eventName, event => {event.preventDefault(); dropZone.classList.add("drag")});
      for (const eventName of ["dragleave", "drop"]) dropZone.addEventListener(eventName, event => {event.preventDefault(); dropZone.classList.remove("drag")});
      dropZone.addEventListener("drop", event => load(event.dataTransfer.files[0]));
    });
  }

  globalThis.KSK_AUTH = {unlock};
})();
