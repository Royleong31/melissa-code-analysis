const form = document.querySelector(".form");

form.addEventListener("submit", (e) => {
  e.preventDefault();

  const input = document.querySelector(".commandInput");
  const output = document.querySelector(".result");
  const command = input.value;

  // port number needs to match the port number in the victim server
  fetch("http://localhost:8080?q=" + command)
    .then((res) => res.json())
    .then((data) => {
      output.innerHTML = data;
    })
    .catch((err) => {
      console.error(err);
    });
});
