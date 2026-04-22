(function () {
  const loadingScreen = document.getElementById("loading-screen");
  const appRoot = document.getElementById("app");
  const threeContainer = document.getElementById("three-container");
  const projectTitle = document.getElementById("project-title");
  const projectStep = document.getElementById("project-step");
  const contactForm = document.getElementById("contact-form");
  const formStatus = document.getElementById("form-status");

  function revealApp() {
    if (!loadingScreen || !appRoot) return;
    loadingScreen.classList.add("hide");
    appRoot.classList.remove("hidden-on-load");
    if (window.gsap) {
      window.gsap.fromTo(appRoot, { opacity: 0 }, { opacity: 1, duration: 0.8, ease: "power2.out" });
    } else {
      appRoot.style.opacity = "1";
    }
  }

  // Hard fallback so loading screen never blocks content.
  window.setTimeout(revealApp, 1800);

  const hasThree = !!window.THREE;
  if (!hasThree || !threeContainer) {
    revealApp();
    if (projectStep) {
      projectStep.textContent = "3D library failed to load. Showing content without 3D objects.";
    }
    setupForm();
    return;
  }

  const THREE = window.THREE;
  const gsap = window.gsap || null;
  const ScrollTrigger = window.ScrollTrigger;
  if (gsap && ScrollTrigger) gsap.registerPlugin(ScrollTrigger);

  const scene = new THREE.Scene();
  scene.background = new THREE.Color("#0a0a0a");

  const camera = new THREE.PerspectiveCamera(50, window.innerWidth / window.innerHeight, 0.1, 100);
  camera.position.set(0, 0.4, 8.8);

  const renderer = new THREE.WebGLRenderer({ antialias: true, alpha: true, powerPreference: "high-performance" });
  renderer.setPixelRatio(Math.min(window.devicePixelRatio, 2));
  renderer.setSize(window.innerWidth, window.innerHeight);
  threeContainer.appendChild(renderer.domElement);

  const ambient = new THREE.AmbientLight("#a3adff", 0.85);
  const key = new THREE.DirectionalLight("#7be9ff", 1.2);
  key.position.set(3, 4, 4);
  const fill = new THREE.DirectionalLight("#c48dff", 0.9);
  fill.position.set(-3, -2, 3);
  scene.add(ambient, key, fill);

  const group = new THREE.Group();
  scene.add(group);

  const definitions = [
    {
      name: "Cube",
      description: "First object: Cube",
      geometry: function () {
        return new THREE.BoxGeometry(1.3, 1.3, 1.3);
      },
      color: "#5cecff",
      x: -3.3,
    },
    {
      name: "Circle",
      description: "Second object: Circle (ring)",
      geometry: function () {
        return new THREE.TorusGeometry(0.95, 0.23, 24, 120);
      },
      color: "#9f80ff",
      x: -1.1,
    },
    {
      name: "Sphere",
      description: "Third object: Sphere",
      geometry: function () {
        return new THREE.SphereGeometry(0.9, 42, 42);
      },
      color: "#7cffb2",
      x: 1.1,
    },
    {
      name: "Cone",
      description: "Fourth object: Cone",
      geometry: function () {
        return new THREE.ConeGeometry(0.82, 1.6, 36);
      },
      color: "#ffd46f",
      x: 3.3,
    },
  ];

  const meshes = definitions.map(function (item, index) {
    const material = new THREE.MeshStandardMaterial({
      color: item.color,
      emissive: item.color,
      emissiveIntensity: 0.2,
      roughness: 0.35,
      metalness: 0.1,
    });

    const mesh = new THREE.Mesh(item.geometry(), material);
    mesh.position.set(item.x, 0, 0);
    mesh.userData.index = index;
    group.add(mesh);
    return mesh;
  });

  if (projectTitle) {
    projectTitle.textContent = "Cube → Circle → Sphere → Cone";
  }
  if (projectStep) {
    projectStep.textContent = "Now showing all objects clearly in sequence.";
  }

  // Simple reveal animation so objects appear one after another.
  meshes.forEach(function (mesh, i) {
    mesh.scale.setScalar(0.001);
    if (gsap) {
      gsap.to(mesh.scale, {
        x: 1,
        y: 1,
        z: 1,
        delay: 0.2 + i * 0.25,
        duration: 0.45,
        ease: "back.out(1.7)",
      });
    } else {
      window.setTimeout(function () {
        mesh.scale.setScalar(1);
      }, 200 + i * 250);
    }
  });

  const mouse = { x: 0, y: 0 };
  const lookAtTarget = new THREE.Vector3(0, 0, 0);

  window.addEventListener("pointermove", function (event) {
    mouse.x = (event.clientX / window.innerWidth) * 2 - 1;
    mouse.y = (event.clientY / window.innerHeight) * 2 - 1;
  });

  window.addEventListener("resize", function () {
    camera.aspect = window.innerWidth / window.innerHeight;
    camera.updateProjectionMatrix();
    renderer.setSize(window.innerWidth, window.innerHeight);
  });

  if (gsap && ScrollTrigger) {
    gsap.utils.toArray(".fade-up").forEach(function (el) {
      gsap.fromTo(
        el,
        { opacity: 0, y: 20 },
        {
          opacity: 1,
          y: 0,
          duration: 0.8,
          ease: "power2.out",
          scrollTrigger: {
            trigger: el,
            start: "top 86%",
          },
        }
      );
    });
  } else {
    document.querySelectorAll(".fade-up").forEach(function (el) {
      el.style.opacity = "1";
      el.style.transform = "translateY(0)";
    });
  }

  const clock = new THREE.Clock();
  function animate() {
    requestAnimationFrame(animate);
    const t = clock.getElapsedTime();

    group.position.y = Math.sin(t * 0.8) * 0.12;

    meshes.forEach(function (mesh, i) {
      mesh.rotation.x += 0.004 + i * 0.0007;
      mesh.rotation.y += 0.007 + i * 0.001;
    });

    group.rotation.y += (mouse.x * 0.2 - group.rotation.y) * 0.04;
    group.rotation.x += (-mouse.y * 0.08 - group.rotation.x) * 0.04;

    camera.position.x += (mouse.x * 0.45 - camera.position.x) * 0.05;
    camera.position.y += (0.4 - mouse.y * 0.22 - camera.position.y) * 0.05;
    camera.lookAt(lookAtTarget);

    renderer.render(scene, camera);
  }

  animate();
  revealApp();
  setupForm();

  function setupForm() {
    if (!contactForm) return;

    contactForm.addEventListener("submit", function (event) {
      event.preventDefault();

      const formData = new FormData(contactForm);
      const payload = {
        name: String(formData.get("name") || ""),
        email: String(formData.get("email") || ""),
        message: String(formData.get("message") || ""),
        at: new Date().toISOString(),
      };

      const key = "portfolio_contact_submissions";
      try {
        const existing = JSON.parse(localStorage.getItem(key) || "[]");
        existing.unshift(payload);
        localStorage.setItem(key, JSON.stringify(existing.slice(0, 20)));
      } catch (_e) {
        // ignore storage errors
      }

      if (formStatus) {
        formStatus.textContent = "Thanks! Message saved locally (demo mode).";
      }
      contactForm.reset();
    });
  }
})();
