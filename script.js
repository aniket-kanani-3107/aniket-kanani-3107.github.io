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

  // Never get stuck on loading screen.
  window.setTimeout(revealApp, 1800);

  const hasThree = !!window.THREE;
  const hasGsap = !!window.gsap;

  if (!hasThree || !hasGsap) {
    revealApp();
    if (projectStep) {
      projectStep.textContent = "3D libraries failed to load. Showing website content without 3D animations.";
    }
    setupForm();
    return;
  }

  const THREE = window.THREE;
  const gsap = window.gsap;
  const ScrollTrigger = window.ScrollTrigger;
  if (ScrollTrigger) gsap.registerPlugin(ScrollTrigger);

  const scene = new THREE.Scene();
  scene.background = new THREE.Color("#0a0a0a");
  scene.fog = new THREE.Fog("#0a0a0a", 8, 16);

  const camera = new THREE.PerspectiveCamera(45, window.innerWidth / window.innerHeight, 0.1, 100);
  camera.position.set(0, 0, 6);

  const renderer = new THREE.WebGLRenderer({ antialias: true, powerPreference: "high-performance" });
  renderer.setPixelRatio(Math.min(window.devicePixelRatio, 2));
  renderer.setSize(window.innerWidth, window.innerHeight);
  threeContainer.appendChild(renderer.domElement);

  const ambient = new THREE.AmbientLight("#9aa7ff", 0.65);
  const key = new THREE.DirectionalLight("#7edbff", 1.3);
  key.position.set(2.8, 3.2, 4.4);
  const rim = new THREE.DirectionalLight("#a275ff", 0.8);
  rim.position.set(-3.2, -2.3, 2.4);
  scene.add(ambient, key, rim);

  const group = new THREE.Group();
  scene.add(group);

  const defs = [
    {
      title: "Neon Cube Study",
      subtitle: "Solid to wireframe to particles",
      geometry: function () { return new THREE.BoxGeometry(1.9, 1.9, 1.9, 14, 14, 14); },
      color: "#57e8ff",
    },
    {
      title: "Signal Sphere",
      subtitle: "Scattered points re-form into smooth volume",
      geometry: function () { return new THREE.SphereGeometry(1.35, 56, 56); },
      color: "#a77dff",
    },
    {
      title: "Orbit Torus",
      subtitle: "Edges tighten into a clean final form",
      geometry: function () { return new THREE.TorusGeometry(1.12, 0.42, 30, 120); },
      color: "#7dffad",
    },
  ];

  const objects = defs.map(function (item) {
    const geometry = item.geometry();
    const basePositions = geometry.attributes.position.array.slice();

    const scatterDirections = new Float32Array(basePositions.length);
    for (let i = 0; i < scatterDirections.length; i += 3) {
      const v = new THREE.Vector3((Math.random() - 0.5) * 2, (Math.random() - 0.5) * 2, (Math.random() - 0.5) * 2).normalize();
      scatterDirections[i] = v.x;
      scatterDirections[i + 1] = v.y;
      scatterDirections[i + 2] = v.z;
    }

    const meshMaterial = new THREE.MeshStandardMaterial({
      color: item.color,
      emissive: item.color,
      emissiveIntensity: 0.18,
      roughness: 0.34,
      metalness: 0.14,
      transparent: true,
      opacity: 0,
      wireframe: false,
    });

    const mesh = new THREE.Mesh(geometry, meshMaterial);

    const pointsGeometry = geometry.clone();
    const pointsMaterial = new THREE.PointsMaterial({
      color: item.color,
      size: 0.03,
      transparent: true,
      opacity: 0,
      sizeAttenuation: true,
      depthWrite: false,
    });

    const points = new THREE.Points(pointsGeometry, pointsMaterial);
    mesh.visible = false;
    points.visible = false;
    group.add(mesh, points);

    return {
      title: item.title,
      subtitle: item.subtitle,
      mesh: mesh,
      points: points,
      basePositions: basePositions,
      scatterDirections: scatterDirections,
    };
  });

  const mouse = { x: 0, y: 0 };
  const cameraTarget = new THREE.Vector3(0, 0, 0);
  let scrollProgress = 0;

  function clamp01(v) {
    return Math.min(1, Math.max(0, v));
  }

  function remap(v, min, max) {
    return clamp01((v - min) / (max - min));
  }

  function setCopy(index) {
    if (projectTitle) projectTitle.textContent = objects[index].title;
    if (projectStep) projectStep.textContent = objects[index].subtitle;
  }

  function updateScatter(obj, amount, shrink) {
    const positions = obj.points.geometry.attributes.position.array;
    for (let i = 0; i < positions.length; i += 3) {
      positions[i] = obj.basePositions[i] * shrink + obj.scatterDirections[i] * amount;
      positions[i + 1] = obj.basePositions[i + 1] * shrink + obj.scatterDirections[i + 1] * amount;
      positions[i + 2] = obj.basePositions[i + 2] * shrink + obj.scatterDirections[i + 2] * amount;
    }
    obj.points.geometry.attributes.position.needsUpdate = true;
  }

  function hideAll() {
    objects.forEach(function (obj) {
      obj.mesh.visible = false;
      obj.points.visible = false;
      obj.mesh.material.opacity = 0;
      obj.mesh.material.wireframe = false;
      obj.points.material.opacity = 0;
      obj.mesh.scale.setScalar(1);
      updateScatter(obj, 0, 1);
    });
  }

  function showSolid(index) {
    hideAll();
    const obj = objects[index];
    obj.mesh.visible = true;
    obj.mesh.material.opacity = 1;
    setCopy(index);
  }

  function transition(fromIndex, toIndex, t) {
    const from = objects[fromIndex];
    const to = objects[toIndex];

    from.mesh.visible = true;
    from.points.visible = true;
    to.mesh.visible = true;
    to.points.visible = true;

    const outFade = remap(t, 0, 0.35);
    const outWire = remap(t, 0.2, 0.45);
    const outPoints = remap(t, 0.35, 0.65);

    const inPoints = remap(t, 0.45, 0.7);
    const inWire = remap(t, 0.65, 0.85);
    const inSolid = remap(t, 0.8, 1);

    from.mesh.material.opacity = 1 - outFade;
    from.mesh.material.wireframe = outWire > 0.08;
    from.mesh.scale.setScalar(1 - outPoints * 0.24);
    from.points.material.opacity = outPoints * (1 - inSolid * 0.8);
    updateScatter(from, outPoints * 2.15, 1 - outPoints * 0.22);

    to.points.material.opacity = inPoints * (1 - inSolid * 0.88);
    updateScatter(to, (1 - inPoints) * 2.35, 0.72 + inPoints * 0.28);
    to.mesh.material.wireframe = inWire > 0.1 && inSolid < 0.92;
    to.mesh.material.opacity = inSolid;
    to.mesh.scale.setScalar(0.82 + inPoints * 0.18);

    setCopy(t < 0.5 ? fromIndex : toIndex);
  }

  function updateSceneFromScroll(progress) {
    const segments = objects.length - 1;
    const scaled = progress * segments;
    const segment = Math.floor(scaled);

    if (progress <= 0) {
      showSolid(0);
      return;
    }
    if (progress >= 0.999) {
      showSolid(objects.length - 1);
      return;
    }

    hideAll();
    const from = Math.min(segment, segments - 1);
    const to = Math.min(from + 1, objects.length - 1);
    transition(from, to, scaled - from);
  }

  showSolid(0);

  window.addEventListener("pointermove", function (event) {
    mouse.x = (event.clientX / window.innerWidth) * 2 - 1;
    mouse.y = (event.clientY / window.innerHeight) * 2 - 1;
  });

  window.addEventListener("resize", function () {
    camera.aspect = window.innerWidth / window.innerHeight;
    camera.updateProjectionMatrix();
    renderer.setSize(window.innerWidth, window.innerHeight);
  });

  if (ScrollTrigger) {
    ScrollTrigger.create({
      trigger: "#portfolio",
      start: "top top",
      end: "bottom bottom",
      scrub: 1,
      onUpdate: function (self) {
        scrollProgress = self.progress;
        updateSceneFromScroll(scrollProgress);
      },
    });

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

    group.position.y = Math.sin(t * 0.8) * 0.1;
    group.rotation.y += (mouse.x * 0.35 + scrollProgress * 0.25 - group.rotation.y) * 0.05;
    group.rotation.x += (-mouse.y * 0.22 - group.rotation.x) * 0.05;

    camera.position.x += (mouse.x * 0.55 + scrollProgress * 0.24 - camera.position.x) * 0.06;
    camera.position.y += (0.12 - mouse.y * 0.28 - scrollProgress * 0.08 - camera.position.y) * 0.06;
    camera.position.z += (6 - scrollProgress * 1.25 - camera.position.z) * 0.06;
    camera.lookAt(cameraTarget);

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
