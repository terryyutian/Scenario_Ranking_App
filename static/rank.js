function updateOrder() {
  const list = document.querySelector("#rank-list");
  const cards = Array.from(list.querySelectorAll(".scenario-card"));
  const orderInput = document.querySelector("#ranking-order");

  orderInput.value = cards.map((card) => card.dataset.scenarioType).join(",");

  cards.forEach((card, index) => {
    const badge = card.querySelector(".rank-badge");
    badge.textContent = `Rank ${index + 1}`;
  });
}

function getDragAfterElement(container, y) {
  const draggableElements = [
    ...container.querySelectorAll(".scenario-card:not(.dragging)"),
  ];

  return draggableElements.reduce(
    (closest, child) => {
      const box = child.getBoundingClientRect();
      const offset = y - box.top - box.height / 2;

      if (offset < 0 && offset > closest.offset) {
        return { offset, element: child };
      }

      return closest;
    },
    { offset: Number.NEGATIVE_INFINITY, element: null },
  ).element;
}

function moveCard(card, direction) {
  if (direction === "up" && card.previousElementSibling) {
    card.parentElement.insertBefore(card, card.previousElementSibling);
  }

  if (direction === "down" && card.nextElementSibling) {
    card.parentElement.insertBefore(card.nextElementSibling, card);
  }

  updateOrder();
}

document.addEventListener("DOMContentLoaded", () => {
  const list = document.querySelector("#rank-list");
  if (!list) {
    return;
  }

  list.addEventListener("dragstart", (event) => {
    const card = event.target.closest(".scenario-card");
    if (!card) {
      return;
    }
    card.classList.add("dragging");
  });

  list.addEventListener("dragend", (event) => {
    const card = event.target.closest(".scenario-card");
    if (!card) {
      return;
    }
    card.classList.remove("dragging");
    updateOrder();
  });

  list.addEventListener("dragover", (event) => {
    event.preventDefault();
    const draggingCard = list.querySelector(".dragging");
    if (!draggingCard) {
      return;
    }

    const afterElement = getDragAfterElement(list, event.clientY);
    if (afterElement === null) {
      list.appendChild(draggingCard);
    } else {
      list.insertBefore(draggingCard, afterElement);
    }
  });

  list.addEventListener("click", (event) => {
    const button = event.target.closest("button");
    if (!button) {
      return;
    }

    const card = button.closest(".scenario-card");
    if (button.classList.contains("move-up")) {
      moveCard(card, "up");
    }

    if (button.classList.contains("move-down")) {
      moveCard(card, "down");
    }
  });

  updateOrder();
});
