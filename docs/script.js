const screenshots = [
  {
    src: "installer-images/Step1.png",
    alt: "Setup Options",
    caption: "Choose the defaults (Install and Create RDP shortcuts) for new installs"
  },
  {
    src: "installer-images/Step2.png",
    alt: "Create RDP User",
    caption: "Create a dedicated RDP user account during installation"
  },
  {
    src: "installer-images/Step3.png",
    alt: "Shortcut Settings",
    caption: "Configure your RDP shortcut settings"
  },
  {
    src: "installer-images/Step4.png",
    alt: "Installation Complete",
    caption: "Installation is complete. Restart if prompted"
  },
  {
    src: "https://github.com/user-attachments/assets/15e7947b-9a11-42b3-b9c7-5f081ce2947f",
    alt: "Desktop Shortcuts",
    caption: "Ready-to-use RDP shortcuts appear on your desktop after setup"
  }
];

const year = document.getElementById("year");
if (year) {
  year.textContent = String(new Date().getFullYear());
}

const lightbox = document.getElementById("lightbox");
const lightboxImage = document.getElementById("lightbox-image");
const lightboxCaption = document.getElementById("lightbox-caption");
const closeLightbox = document.getElementById("close-lightbox");
const prevShot = document.getElementById("prev-shot");
const nextShot = document.getElementById("next-shot");
const shotButtons = document.querySelectorAll(".shot");

let activeShot = 0;

const renderLightbox = () => {
  if (!lightboxImage || !lightboxCaption) {
    return;
  }

  const shot = screenshots[activeShot];
  lightboxImage.src = shot.src;
  lightboxImage.alt = shot.alt;
  lightboxCaption.textContent = "Step " + (activeShot + 1) + " of " + screenshots.length + " - " + shot.caption;
};

const openLightbox = (index) => {
  if (!lightbox) {
    return;
  }

  activeShot = index;
  renderLightbox();
  lightbox.hidden = false;
  document.body.style.overflow = "hidden";
};

const closeLightboxFn = () => {
  if (!lightbox) {
    return;
  }

  lightbox.hidden = true;
  document.body.style.overflow = "";
};

shotButtons.forEach((button) => {
  button.addEventListener("click", () => {
    const index = Number(button.getAttribute("data-shot"));
    if (!Number.isNaN(index)) {
      openLightbox(index);
    }
  });
});

if (closeLightbox && lightbox) {
  closeLightbox.addEventListener("click", () => {
    closeLightboxFn();
  });
}

if (prevShot && lightbox) {
  prevShot.addEventListener("click", () => {
    activeShot = (activeShot - 1 + screenshots.length) % screenshots.length;
    renderLightbox();
  });
}

if (nextShot && lightbox) {
  nextShot.addEventListener("click", () => {
    activeShot = (activeShot + 1) % screenshots.length;
    renderLightbox();
  });
}

if (lightbox) {
  lightbox.addEventListener("click", (event) => {
    if (event.target === lightbox) {
      closeLightboxFn();
    }
  });
}

document.addEventListener("keydown", (event) => {
  if (!lightbox || lightbox.hidden) {
    return;
  }

  if (event.key === "Escape") {
    closeLightboxFn();
  }

  if (event.key === "ArrowLeft") {
    activeShot = (activeShot - 1 + screenshots.length) % screenshots.length;
    renderLightbox();
  }

  if (event.key === "ArrowRight") {
    activeShot = (activeShot + 1) % screenshots.length;
    renderLightbox();
  }
});

const setupScrollReveal = () => {
  const prefersReducedMotion = window.matchMedia("(prefers-reduced-motion: reduce)").matches;
  if (prefersReducedMotion) {
    return;
  }

  const selectors = [
    ".hero-content > *",
    ".section h2",
    ".section-subtitle",
    ".cards-grid .card",
    ".screenshots-grid .shot",
    ".steps li",
  ];

  const revealTargets = document.querySelectorAll(selectors.join(","));
  revealTargets.forEach((element) => {
    element.classList.add("reveal-on-scroll");
  });

  const staggerGroups = [
    ".cards-grid .card",
    ".screenshots-grid .shot",
    ".steps li"
  ];

  staggerGroups.forEach((groupSelector) => {
    const groupItems = document.querySelectorAll(groupSelector);
    groupItems.forEach((item, index) => {
      item.style.setProperty("--reveal-delay", String(index * 90) + "ms");
    });
  });

  const observer = new IntersectionObserver(
    (entries) => {
      entries.forEach((entry) => {
        if (entry.isIntersecting) {
          entry.target.classList.add("is-visible");
          observer.unobserve(entry.target);
        }
      });
    },
    {
      threshold: 0.14,
      rootMargin: "0px 0px -10% 0px"
    }
  );

  revealTargets.forEach((element) => {
    observer.observe(element);
  });
};

setupScrollReveal();