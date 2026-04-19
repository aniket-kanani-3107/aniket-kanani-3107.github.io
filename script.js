import * as THREE from "https://unpkg.com/three@0.164.1/build/three.module.js";
import { EffectComposer } from "https://unpkg.com/three@0.164.1/examples/jsm/postprocessing/EffectComposer.js";
import { RenderPass } from "https://unpkg.com/three@0.164.1/examples/jsm/postprocessing/RenderPass.js";
import { UnrealBloomPass } from "https://unpkg.com/three@0.164.1/examples/jsm/postprocessing/UnrealBloomPass.js";
import { gsap } from "https://esm.sh/gsap@3.12.5";
import { ScrollTrigger } from "https://esm.sh/gsap@3.12.5/ScrollTrigger";

gsap.registerPlugin(ScrollTrigger);

const threeContainer = document.querySelector("#three-container");
const loadingScreen = document.querySelector("#loading-screen");
const appRoot = document.querySelector("#app");
const projectTitle = document.querySelector("#project-title");
const projectStep = document.querySelector("#project-step");
const form = document.querySelector(".contact-form");

const scene = new THREE.Scene();
scene.background = new THREE.Color("#0a0a0a");
scene.fog = new THREE.Fog("#0a0a0a", 8, 16);

const camera = new THREE.PerspectiveCamera(
  45,
  window.innerWidth / window.innerHeight,
  0.1,
  100
);
camera.position.set(0, 0, 6);

const renderer = new THREE.WebGLRenderer({ antialias: true, powerPreference: "high-performance" });
renderer.setPixelRatio(Math.min(window.devicePixelRatio, 2));
renderer.setSize(window.innerWidth, window.innerHeight);
threeContainer.appendChild(renderer.domElement);

// Post-processing for subtle glow around bright edges.
const composer = new EffectComposer(renderer);
composer.addPass(new RenderPass(scene, camera));
composer.addPass(new UnrealBloomPass(new THREE.Vector2(window.innerWidth, window.innerHeight), 0.7, 0.5, 0.8));

// Realistic lighting setup: key + fill + ambient.
const ambientLight = new THREE.AmbientLight("#9aa7ff", 0.6);
const keyLight = new THREE.DirectionalLight("#7edbff", 1.35);
keyLight.position.set(2.8, 3.2, 4.4);
const rimLight = new THREE.DirectionalLight("#9d70ff", 0.8);
rimLight.position.set(-3.5, -2.1, 2.3);
scene.add(ambientLight, keyLight, rimLight);

const portfolioGroup = new THREE.Group();
scene.add(portfolioGroup);

const projectDefinitions = [
  {
    title: "Neon Cube Study",
    subtitle: "Geometric balance and rhythm",
    geometry: () => new THREE.BoxGeometry(1.9, 1.9, 1.9, 14, 14, 14),
    color: "#57e8ff",
  },
  {
    title: "Signal Sphere",
    subtitle: "Organic pulse and smooth surfaces",
    geometry: () => new THREE.SphereGeometry(1.36, 58, 58),
    color: "#a77dff",
  },
  {
    title: "Orbit Torus",
    subtitle: "Looping systems and motion depth",
    geometry: () => new THREE.TorusGeometry(1.12, 0.42, 32, 140),
    color: "#78ffa8",
  },
];

const objects = projectDefinitions.map((item, index) => {
  const geometry = item.geometry();
  const positionAttr = geometry.attributes.position;
  const basePositions = positionAttr.array.slice();

  // Random direction per vertex for the scatter transition stage.
  const scatterDirections = new Float32Array(basePositions.length);
  for (let i = 0; i < scatterDirections.length; i += 3) {
    const vec = new THREE.Vector3(
      (Math.random() - 0.5) * 2,
      (Math.random() - 0.5) * 2,
      (Math.random() - 0.5) * 2
    ).normalize();
    scatterDirections[i] = vec.x;
    scatterDirections[i + 1] = vec.y;
    scatterDirections[i + 2] = vec.z;
  }

  const meshMaterial = new THREE.MeshStandardMaterial({
    color: item.color,
    emissive: item.color,
    emissiveIntensity: 0.2,
    roughness: 0.32,
    metalness: 0.12,
    transparent: true,
    opacity: 0,
    wireframe: false,
  });

  const mesh = new THREE.Mesh(geometry, meshMaterial);

  const pointsGeometry = geometry.clone();
  const pointsMaterial = new THREE.PointsMaterial({
    color: item.color,
    size: 0.03,
    sizeAttenuation: true,
    transparent: true,
    opacity: 0,
    depthWrite: true,
  });
  const points = new THREE.Points(pointsGeometry, pointsMaterial);

  mesh.visible = false;
  points.visible = false;
  portfolioGroup.add(mesh, points);

  return {
    index,
    title: item.title,
    subtitle: item.subtitle,
    mesh,
    points,
    basePositions,
    scatterDirections,
  };
});

let portfolioProgress = 0;
const mouse = { x: 0, y: 0 };
const cameraTarget = new THREE.Vector3(0, 0, 0);

const clamp01 = (v) => Math.min(1, Math.max(0, v));
const remap = (v, min, max) => clamp01((v - min) / (max - min));

function setProjectCopy(index) {
  projectTitle.textContent = objects[index].title;
  projectStep.textContent = objects[index].subtitle;
}

// Update particle positions by moving each vertex along its random direction.
function updatePointsScatter(item, amount = 0, shrink = 1) {
  const positions = item.points.geometry.attributes.position.array;

  for (let i = 0; i < positions.length; i += 3) {
    positions[i] = item.basePositions[i] * shrink + item.scatterDirections[i] * amount;
    positions[i + 1] = item.basePositions[i + 1] * shrink + item.scatterDirections[i + 1] * amount;
    positions[i + 2] = item.basePositions[i + 2] * shrink + item.scatterDirections[i + 2] * amount;
  }

  item.points.geometry.attributes.position.needsUpdate = true;
}

function hideAllObjects() {
  for (const item of objects) {
    item.mesh.visible = false;
    item.points.visible = false;
    item.mesh.material.opacity = 0;
    item.mesh.material.wireframe = false;
    item.points.material.opacity = 0;
    item.mesh.scale.setScalar(1);
    updatePointsScatter(item, 0, 1);
  }
}

function showSolidObject(index) {
  hideAllObjects();
  const target = objects[index];
  target.mesh.visible = true;
  target.mesh.material.opacity = 1;
  target.mesh.material.wireframe = false;
  target.points.visible = false;
  setProjectCopy(index);
}

// Complete disappear/appear choreography between two objects.
function animateTransition(fromIndex, toIndex, t) {
  const from = objects[fromIndex];
  const to = objects[toIndex];

  from.mesh.visible = true;
  from.points.visible = true;
  to.mesh.visible = true;
  to.points.visible = true;

  const fadeOut = remap(t, 0, 0.35);
  const fromWire = remap(t, 0.2, 0.4);
  const fromPoints = remap(t, 0.35, 0.6);

  const toStart = remap(t, 0.45, 0.65);
  const toWire = remap(t, 0.65, 0.85);
  const toSolid = remap(t, 0.8, 1);

  // FROM object: solid -> wireframe -> points -> scattered/shrunk.
  from.mesh.material.opacity = 1 - fadeOut;
  from.mesh.material.wireframe = fromWire > 0.05;
  from.mesh.scale.setScalar(1 - fromPoints * 0.25);

  const fromScatter = fromPoints * 2.1;
  from.points.material.opacity = fromPoints * (1 - toSolid * 0.75);
  updatePointsScatter(from, fromScatter, 1 - fromPoints * 0.22);

  // TO object: scattered points -> edge/wire phase -> solid mesh.
  const toScatter = (1 - toStart) * 2.5;
  to.points.material.opacity = toStart * (1 - toSolid * 0.86);
  updatePointsScatter(to, toScatter, 0.7 + toStart * 0.3);

  to.mesh.material.wireframe = toWire > 0.1 && toSolid < 0.92;
  to.mesh.material.opacity = toSolid;
  to.mesh.scale.setScalar(0.82 + toStart * 0.18);

  if (t < 0.5) {
    setProjectCopy(fromIndex);
  } else {
    setProjectCopy(toIndex);
  }
}

function updatePortfolioState(progress) {
  const segments = objects.length - 1;
  const scaled = progress * segments;
  const segment = Math.floor(scaled);

  if (progress <= 0) {
    showSolidObject(0);
    return;
  }

  if (progress >= 0.999) {
    showSolidObject(objects.length - 1);
    return;
  }

  hideAllObjects();
  const fromIndex = Math.min(segment, segments - 1);
  const toIndex = Math.min(fromIndex + 1, objects.length - 1);
  const localProgress = scaled - fromIndex;
  animateTransition(fromIndex, toIndex, localProgress);
}

showSolidObject(0);

// Smoothly map mouse movement for parallax and tiny rotational interaction.
window.addEventListener("pointermove", (event) => {
  mouse.x = (event.clientX / window.innerWidth) * 2 - 1;
  mouse.y = (event.clientY / window.innerHeight) * 2 - 1;
});

window.addEventListener("resize", () => {
  camera.aspect = window.innerWidth / window.innerHeight;
  camera.updateProjectionMatrix();
  renderer.setSize(window.innerWidth, window.innerHeight);
  composer.setSize(window.innerWidth, window.innerHeight);
});

ScrollTrigger.create({
  trigger: "#portfolio",
  start: "top top",
  end: "bottom bottom",
  scrub: 1,
  onUpdate: (self) => {
    portfolioProgress = self.progress;
    updatePortfolioState(portfolioProgress);
  },
});

// Fade-up entrance for section cards.
gsap.utils.toArray(".fade-up").forEach((el) => {
  gsap.fromTo(
    el,
    { opacity: 0, y: 22 },
    {
      opacity: 1,
      y: 0,
      duration: 1,
      ease: "power2.out",
      scrollTrigger: {
        trigger: el,
        start: "top 86%",
      },
    }
  );
});

const clock = new THREE.Clock();
function renderLoop() {
  requestAnimationFrame(renderLoop);

  const t = clock.getElapsedTime();
  const motionX = mouse.x * 0.35;
  const motionY = -mouse.y * 0.22;

  // Floating + slow spin keeps the scene alive even while not scrolling.
  portfolioGroup.position.y = Math.sin(t * 0.8) * 0.1;
  portfolioGroup.rotation.y += (motionX + portfolioProgress * 0.22 - portfolioGroup.rotation.y) * 0.04;
  portfolioGroup.rotation.x += (motionY - portfolioGroup.rotation.x) * 0.05;

  camera.position.x += (mouse.x * 0.55 + portfolioProgress * 0.25 - camera.position.x) * 0.06;
  camera.position.y += (0.15 + -mouse.y * 0.3 - portfolioProgress * 0.1 - camera.position.y) * 0.06;
  camera.position.z += (6 - portfolioProgress * 1.25 - camera.position.z) * 0.06;
  camera.lookAt(cameraTarget);

  composer.render();
}

renderLoop();

window.addEventListener("load", () => {
  // Startup reveal: loading overlay fades out and content fades in.
  gsap.delayedCall(0.45, () => {
    loadingScreen.classList.add("hide");
    appRoot.classList.remove("hidden-on-load");
    gsap.fromTo(appRoot, { opacity: 0 }, { opacity: 1, duration: 0.9, ease: "power2.out" });
  });
});

if (form) {
  form.addEventListener("submit", (event) => {
    event.preventDefault();
    const button = form.querySelector("button");
    const originalText = button.textContent;
    button.textContent = "Message Sent ✓";
    button.disabled = true;

    setTimeout(() => {
      button.textContent = originalText;
      button.disabled = false;
      form.reset();
    }, 1800);
  });
}
