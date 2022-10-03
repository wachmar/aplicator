"use strict";

const otherSourceDiv = document.getElementById("other_source_div");
const otherSourceInput = document.getElementById("other_source");
const lSelect = document.getElementById("source");

lSelect.addEventListener("change", function () {
  if (Number(lSelect.value) === 99) {
    otherSourceDiv.classList.remove("d-none");
    otherSourceInput.required = true;
  } else {
    otherSourceDiv.classList.add("d-none");
    otherSourceInput.required = false;
  }
});

const hrPersonNameInput = document.getElementById("hr_person_name");
const radiosHrPerson = document.getElementsByName("gender");
for (const radio of radiosHrPerson) {
  radio.onclick = function (e) {
    console.log(e.target.value);
    if (e.target.value === "not_known") {
      hrPersonNameInput.required = false;
    } else {
      hrPersonNameInput.required = true;
    }
  };
}
