export function createNameBoxElement(controller) {
  const root = document.createElement("div");
  root.className = "name-box";

  const input = document.createElement("input");
  input.type = "text";
  input.className = "name-box-input";
  input.value = controller.getDisplayValue();

  input.addEventListener("focus", () => {
    input.select();
  });

  input.addEventListener("keydown", (e) => {
    if (e.key === "Enter") {
      e.preventDefault();
      controller.submit(input.value);
      input.value = controller.getDisplayValue();
      input.blur();
    } else if (e.key === "Escape") {
      e.preventDefault();
      input.value = controller.getDisplayValue();
      input.blur();
    }
  });

  root.append(input);

  root.refresh = () => {
    input.value = controller.getDisplayValue();
  };

  return root;
}
