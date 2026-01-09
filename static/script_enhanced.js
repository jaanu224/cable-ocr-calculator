// =============== BASIC DOM HANDLES ===============
const pdfInput = document.getElementById("pdfFile");
const btnExtract = document.getElementById("btnExtract");
// extractStatus element removed from HTML


const conductorForm = document.getElementById("conductorForm");
const sheathForm = document.getElementById("sheathForm");
const resultBox = document.getElementById("resultBox");
const resultText = document.getElementById("resultText");
const btnReset = document.getElementById("btnReset");

// Download buttons
const btnDownloadConductor = document.getElementById("btnDownloadConductor");
const btnDownloadSheath = document.getElementById("btnDownloadSheath");

// Store calculation data
let conductorData = null;
let sheathData = null;
let conductorCalculated = false;
let sheathCalculated = false;

// PDF preview button
const btnViewPdf = document.getElementById("btnViewPdf");
let uploadedPdfUrl = null;

// Rated-voltage dropdown (optional)
const ratedVoltageSelect = document.getElementById("ratedVoltageSelect");
const ratedVoltageHelp = document.getElementById("ratedVoltageHelp");

// Dropdowns for insulation & outer sheath
const insulationSelect = document.getElementById("insulationMaterial");
const outerSheathSelect = document.getElementById("outerSheathMaterial");

// =============== CONSTANT TABLES ===============

// Sheaths, screens, armour (Table I)
const TABLE_I_SHEATHS = {
  lead: { K: 41, beta: 230, sigmaC: 1.45e6, rho20: 21.4e-8 },
  steel: { K: 78, beta: 202, sigmaC: 3.8e6, rho20: 13.8e-8 },
  bronze: { K: 180, beta: 313, sigmaC: 3.4e6, rho20: 3.5e-8 },
  aluminium: { K: 148, beta: 228, sigmaC: 2.5e6, rho20: 2.84e-8 }
};

// Conductors (Table I)
const TABLE_I_CONDUCTORS = {
  copper: { K: 226, beta: 234.5, sigmaC: 3.45e6, rho20: 1.7241e-8 },
  aluminium: { K: 148, beta: 228, sigmaC: 2.5e6, rho20: 2.8264e-8 }
};

// Thermal constants (ρ, σ)
const THERMAL_CONSTANTS = {
  insulating: {
    "impregnated-paper-solid": { rho: 6.0, sigma: 2.0e6 },
    "impregnated-paper-oil-filled": { rho: 5.0, sigma: 2.0e6 },
    oil: { rho: 7.0, sigma: 1.7e6 },
    PE: { rho: 3.5, sigma: 2.4e6 },
    XLPE: { rho: 3.5, sigma: 2.4e6 },
    PVC: {
      "<=3kV": { rho: 5.0, sigma: 1.7e6 },
      ">3kV": { rho: 6.0, sigma: 1.7e6 }
    },
    EPR: {
      "<=3kV": { rho: 3.5, sigma: 2.0e6 },
      ">3kV": { rho: 5.0, sigma: 2.0e6 }
    },
    "butyl-rubber": { rho: 5.0, sigma: 2.0e6 },
    "natural-rubber": { rho: 5.0, sigma: 2.0e6 }
  },
  protective: {
    "compounded-jute": { rho: 6.0, sigma: 2.0e6 },
    "rubber-sandwich": { rho: 6.0, sigma: 2.0e6 },
    polychloroprene: { rho: 5.5, sigma: 2.0e6 },
    PVC: {
      "<=35kV": { rho: 5.0, sigma: 1.7e6 },
      ">35kV": { rho: 6.0, sigma: 1.7e6 }
    },
    "PVC-bitumen": { rho: 6.0, sigma: 1.7e6 },
    PE: { rho: 3.5, sigma: 2.4e6 }
  }
};

// Thermal contact factor F
const THERMAL_CONTACT_FACTOR = {
  default: 0.7,
  "oil-filled": 1.0
};

// =============== INITIAL DEFAULTS ===============

// Default insulation = XLPE if nothing is selected
if (insulationSelect && !insulationSelect.value) {
  insulationSelect.value = "XLPE";
}

// Default outer sheath = PE if nothing is selected
if (outerSheathSelect && !outerSheathSelect.value) {
  outerSheathSelect.value = "PE";
}

// =============== UTILS ===============
function setValue(id, value) {
  const el = document.getElementById(id);
  if (!el) return;
  if (value === null || value === undefined) return;
  el.value = value;
}

function showResult(html, isError = false) {
  console.log("showResult called with html length:", html.length);
  console.log("HTML preview:", html.substring(0, 150));
  resultText.innerHTML = html;
  resultText.className = isError ? 'results-box error' : 'results-box';
  resultBox.style.display = "block";
  console.log("Result box display set to block");
  console.log("resultText.innerHTML is now:", resultText.innerHTML.substring(0, 150));

  // Force scroll after a short delay to ensure rendering
  setTimeout(() => {
    if (resultBox && typeof resultBox.scrollIntoView === "function") {
      resultBox.scrollIntoView({ behavior: "smooth", block: "center" });
      console.log("Scrolled to result box");
    }
  }, 100);
}

function showNotification(message, type = 'info') {
  // Simple notification system
  const notification = document.createElement('div');
  notification.style.cssText = `
    position: fixed;
    top: 20px;
    right: 20px;
    padding: 1rem 1.5rem;
    border-radius: 10px;
    color: white;
    font-weight: 600;
    z-index: 9999;
    animation: slideIn 0.3s ease;
    box-shadow: 0 4px 12px rgba(0,0,0,0.15);
  `;
  
  const colors = {
    success: '#10b981',
    error: '#ef4444',
    warning: '#f59e0b',
    info: '#2563eb'
  };
  
  notification.style.background = colors[type] || colors.info;
  notification.textContent = message;
  
  document.body.appendChild(notification);
  
  setTimeout(() => {
    notification.style.animation = 'slideOut 0.3s ease';
    setTimeout(() => notification.remove(), 300);
  }, 3000);
}

function validateVoltageTime(voltageKv, t) {
  if (voltageKv <= 0 || t <= 0) {
    showNotification("Voltage and time must be positive.", 'error');
    return false;
  }
  if (voltageKv > 400) {
    showNotification("Voltage must not be greater than 400 kV.", 'error');
    return false;
  }
  if (t > 10) {
    showNotification("Time must not be greater than 10 seconds.", 'error');
    return false;
  }
  return true;
}

// =============== Rated Voltage dropdown helper ===============
function populateRatedVoltages(list, headerVoltage) {
  if (!ratedVoltageSelect || !ratedVoltageHelp) return;
  
  const voltageDropdownBtn = document.getElementById("voltageDropdownBtn");

  ratedVoltageSelect.innerHTML = "";
  if (!list || !list.length) {
    ratedVoltageSelect.style.display = "none";
    ratedVoltageHelp.style.display = "none";
    if (voltageDropdownBtn) voltageDropdownBtn.style.display = "none";
    return;
  }

  ratedVoltageSelect.style.display = "block";
  ratedVoltageHelp.style.display = "block";
  if (voltageDropdownBtn) {
    voltageDropdownBtn.style.display = "block";
    console.log("Showing voltage dropdown button");
  }

  const placeholder = document.createElement("option");
  placeholder.value = "";
  placeholder.textContent = "Select rated voltage";
  ratedVoltageSelect.appendChild(placeholder);

  list.forEach((v) => {
    const opt = document.createElement("option");
    opt.value = v;
    opt.textContent = `${v} kV`;
    ratedVoltageSelect.appendChild(opt);
  });

  // Auto-select header voltage if it is in the list
  if (headerVoltage != null) {
    const match = list.find((v) => v === headerVoltage);
    if (match !== undefined) {
      ratedVoltageSelect.value = String(match);
    }
  }
}

if (ratedVoltageSelect) {
  ratedVoltageSelect.addEventListener("change", () => {
    const val = parseFloat(ratedVoltageSelect.value);
    console.log("Dropdown changed to:", val);
    if (!isNaN(val)) {
      setValue("voltageKv", val);
      setValue("sheathVoltageKv", val);
      
      // Close the dropdown immediately
      ratedVoltageSelect.size = 1;
      ratedVoltageSelect.blur();
      console.log("Dropdown closed after selection");
    }
  });
  
  // Close dropdown when clicking outside
  document.addEventListener("click", (e) => {
    if (!ratedVoltageSelect.contains(e.target) && !voltageDropdownBtn.contains(e.target)) {
      ratedVoltageSelect.size = 1;
    }
  });
}

// Add click handler for dropdown button
const voltageDropdownBtn = document.getElementById("voltageDropdownBtn");
if (voltageDropdownBtn) {
  voltageDropdownBtn.addEventListener("click", (e) => {
    e.preventDefault();
    e.stopPropagation();
    console.log("Arrow button clicked!");
    
    if (ratedVoltageSelect && ratedVoltageSelect.style.display !== "none") {
      console.log("Opening voltage dropdown...");
      
      // Show dropdown by setting size to number of options
      const optionCount = ratedVoltageSelect.options.length;
      if (optionCount > 1) {
        // Toggle dropdown - if already open, close it
        if (ratedVoltageSelect.size > 1) {
          ratedVoltageSelect.size = 1;
          console.log("Dropdown closed");
        } else {
          ratedVoltageSelect.size = Math.min(optionCount, 6); // Show max 6 options
          ratedVoltageSelect.focus();
          console.log("Dropdown opened with", optionCount, "options");
        }
      }
    } else {
      console.log("Dropdown not available or hidden");
    }
  });
}

// Add click handler for conductor mode dropdown button
const conductorModeDropdownBtn = document.getElementById("conductorModeDropdownBtn");
const conductorModeSelect = document.getElementById("conductorMode");
if (conductorModeDropdownBtn && conductorModeSelect) {
  conductorModeDropdownBtn.addEventListener("click", (e) => {
    e.preventDefault();
    e.stopPropagation();
    console.log("Conductor mode arrow button clicked!");
    
    // Show dropdown by setting size to number of options
    const optionCount = conductorModeSelect.options.length;
    if (optionCount > 1) {
      // Toggle dropdown - if already open, close it
      if (conductorModeSelect.size > 1) {
        conductorModeSelect.size = 1;
        console.log("Conductor mode dropdown closed");
      } else {
        conductorModeSelect.size = Math.min(optionCount, 6); // Show max 6 options
        conductorModeSelect.focus();
        console.log("Conductor mode dropdown opened with", optionCount, "options");
      }
    }
  });
  
  // Close dropdown when selection is made
  conductorModeSelect.addEventListener("change", () => {
    conductorModeSelect.size = 1;
    conductorModeSelect.blur();
    console.log("Conductor mode dropdown closed after selection");
  });
  
  // Close dropdown when clicking outside
  document.addEventListener("click", (e) => {
    if (!conductorModeSelect.contains(e.target) && !conductorModeDropdownBtn.contains(e.target)) {
      conductorModeSelect.size = 1;
    }
  });
}

// =============== PDF PREVIEW WIRING ===============
pdfInput.addEventListener("change", () => {
  const file = pdfInput.files[0];
  if (!file) {
    uploadedPdfUrl = null;
    if (btnViewPdf) btnViewPdf.style.display = "none";
    return;
  }

  if (uploadedPdfUrl) {
    URL.revokeObjectURL(uploadedPdfUrl);
  }
  uploadedPdfUrl = URL.createObjectURL(file);

  if (btnViewPdf) {
    btnViewPdf.style.display = "inline-block";
  }
  
  showNotification("PDF loaded successfully!", 'success');
});

if (btnViewPdf) {
  btnViewPdf.addEventListener("click", () => {
    if (uploadedPdfUrl) {
      window.open(uploadedPdfUrl, "_blank");
    }
  });
}

// =============== STEP 1: PDF OCR & EXTRACTION ===============
btnExtract.addEventListener("click", async () => {
  const file = pdfInput.files[0];
  if (!file) {
    showNotification("Please choose a PDF file first.", 'warning');
    return;
  }

  // Show extracting notification
  showNotification("Extracting data from PDF...", 'info');
  resultBox.style.display = "none";

  const fd = new FormData();
  fd.append("file", file);

  try {
    const res = await fetch("/api/extract", { method: "POST", body: fd });
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || "Extraction failed");

    console.log("Extraction data from server:", data);

    // Choose a finalVoltage to use:
    let finalVoltage = data.voltageKv;
    if ((finalVoltage == null || isNaN(finalVoltage)) &&
        data.ratedVoltages && data.ratedVoltages.length) {
      finalVoltage = data.ratedVoltages[data.ratedVoltages.length - 1];
    }

    // Voltage / current / time
    setValue("voltageKv", finalVoltage);
    
    // Handle short circuit current with red box if not found
    if (data.sccKa) {
      setValue("sccKa", data.sccKa);
      setValue("sheathSccKa", data.sccKa);
      // Remove any previous error styling
      document.getElementById("sccKa").style.borderColor = "";
      document.getElementById("sheathSccKa").style.borderColor = "";
    } else {
      // Show red box and message when no short circuit current found
      setValue("sccKa", "");
      setValue("sheathSccKa", "");
      
      // Add red border to indicate missing value
      document.getElementById("sccKa").style.borderColor = "#ef4444";
      document.getElementById("sccKa").style.borderWidth = "2px";
      document.getElementById("sccKa").placeholder = "⚠️ Not found in PDF";
      
      document.getElementById("sheathSccKa").style.borderColor = "#ef4444";
      document.getElementById("sheathSccKa").style.borderWidth = "2px";
      document.getElementById("sheathSccKa").placeholder = "⚠️ Not found in PDF";
      
      // Show notification message
      showNotification("⚠️ Short circuit current not found in PDF. Please enter manually.", 'warning');
    }
    
    setValue("timeSec", data.timeSec ?? 1);

    setValue("sheathVoltageKv", finalVoltage);
    setValue("sheathTimeSec", data.timeSec ?? 1);

    // Conductor area (if extracted)
    if (data.conductorArea) {
      setValue("givenConductorArea", data.conductorArea);
    }

    // Sheath dimensions (if extracted with improved OCR)
    console.log("Sheath dimensions from extraction:", {
      outerD: data.sheathOuterD,
      innerD: data.sheathInnerD,
      thickness: data.sheathThickness
    });
    
    if (data.sheathOuterD) {
      console.log("Setting sheathOuterD to:", data.sheathOuterD);
      setValue("sheathOuterD", data.sheathOuterD);
    }
    if (data.sheathInnerD) {
      console.log("Setting sheathInnerD to:", data.sheathInnerD);
      setValue("sheathInnerD", data.sheathInnerD);
    }
    if (data.sheathThickness) {
      console.log("Setting sheathThickness to:", data.sheathThickness);
      setValue("sheathThickness", data.sheathThickness);
    }
    // Trigger geometry update to calculate area
    if (data.sheathOuterD && data.sheathInnerD) {
      console.log("Triggering updateSheathGeometry");
      updateSheathGeometry();
    }

    // Conductor material
    const condMat = data.conductorMaterial || data.material || "";
    if (condMat) {
      document.getElementById("material").value = condMat;
    }

    // Sheath material
    let sheathMat = data.sheathMaterial || "";
    if (!sheathMat && condMat) {
      const lower = condMat.toLowerCase();
      if (lower.includes("al")) sheathMat = "aluminium";
      else sheathMat = "aluminium";
    }
    if (sheathMat) {
      const sheathSelect = document.getElementById("sheathMaterial");
      if (sheathSelect) sheathSelect.value = sheathMat;
    }

    // Insulation / outer sheath
    if (data.insulationMaterial && insulationSelect) {
      insulationSelect.value = data.insulationMaterial;
    }
    if (data.outerSheathMaterial && outerSheathSelect) {
      outerSheathSelect.value = data.outerSheathMaterial;
    }

    // K & β
    setValue("kValue", data.kValue);
    setValue("beta", data.beta);

    // Rated voltages dropdown
    populateRatedVoltages(data.ratedVoltages, data.voltageKv);

    // Debug text removed - keeping extraction logic only

    showNotification("Data extracted successfully!", 'success');
  } catch (err) {
    console.error(err);
    showNotification("Could not extract data from PDF: " + err.message, 'error');
  }
});

// =============== STEP 2: CONDUCTOR CALCULATION ===============
function calculateConductorAreaFromCurrent() {
  const voltageKv = parseFloat(document.getElementById("voltageKv").value);
  const I_AD_kA = parseFloat(document.getElementById("sccKa").value);
  const t = parseFloat(document.getElementById("timeSec").value);
  const K = parseFloat(document.getElementById("kValue").value);
  const beta = parseFloat(document.getElementById("beta").value);
  const theta_i = 90;
  const theta_f = 250;

  if ([voltageKv, I_AD_kA, t, K, beta].some((v) => isNaN(v))) {
    showNotification("Please fill all conductor inputs.", 'warning');
    return null;
  }
  if (!validateVoltageTime(voltageKv, t)) return null;
  if (I_AD_kA <= 0 || K <= 0) {
    showNotification("Current and K must be positive.", 'error');
    return null;
  }

  const lnTerm = Math.log((theta_f + beta) / (theta_i + beta));
  if (lnTerm <= 0) {
    showNotification("Invalid temperature / beta combination.", 'error');
    return null;
  }

  const I_AD_A = I_AD_kA * 1000;
  const S_sq = (I_AD_A ** 2 * t) / (K ** 2 * lnTerm);
  if (S_sq <= 0) {
    showNotification("Calculated conductor area is not valid.", 'error');
    return null;
  }

  return Math.sqrt(S_sq);
}

function calculateConductorCurrentFromArea() {
  const t = parseFloat(document.getElementById("timeSec").value);
  const K = parseFloat(document.getElementById("kValue").value);
  const beta = parseFloat(document.getElementById("beta").value);
  const theta_i = 90;
  const theta_f = 250;
  const S_given = parseFloat(
    document.getElementById("givenConductorArea").value
  );

  if ([t, K, beta, S_given].some((v) => isNaN(v))) {
    showNotification("Please fill time, K, β and given area.", 'warning');
    return null;
  }
  if (t <= 0) {
    showNotification("Time must be positive.", 'error');
    return null;
  }
  if (S_given <= 0 || K <= 0) {
    showNotification("Area and K must be positive.", 'error');
    return null;
  }

  const lnTerm = Math.log((theta_f + beta) / (theta_i + beta));
  if (lnTerm <= 0) {
    showNotification("Invalid temperature / beta combination.", 'error');
    return null;
  }

  const I_AD_A = K * S_given * Math.sqrt(lnTerm / t);
  return I_AD_A / 1000;
}

conductorForm.addEventListener("submit", (e) => {
  e.preventDefault();
  console.log("=== CONDUCTOR FORM SUBMITTED ===");
  // Don't hide result box - just update it
  
  const mode = document.getElementById("conductorMode").value;
  const S_given_str = document.getElementById("givenConductorArea").value;
  let html = "<h6 style='color: #2563eb;'>Conductor Calculation Results</h6><hr>";

  const voltageKv = parseFloat(document.getElementById("voltageKv").value);
  const I_AD_kA = parseFloat(document.getElementById("sccKa").value);
  const t = parseFloat(document.getElementById("timeSec").value);
  const K = parseFloat(document.getElementById("kValue").value);
  const beta = parseFloat(document.getElementById("beta").value);
  const material = document.getElementById("material").value;
  const insulation = document.getElementById("insulationMaterial").value;
  const outerSheath = document.getElementById("outerSheathMaterial").value;
  const theta_i = 90;
  const theta_f = 250;

  if (mode === "area-from-current") {
    const S_required = calculateConductorAreaFromCurrent();
    if (S_required == null) return;

    html += `<p><strong>Required cross-sectional area S:</strong> <span style="font-size: 1.2rem; color: #2563eb;">${S_required.toFixed(2)} mm²</span></p>`;

    const S_given = S_given_str !== "" ? parseFloat(S_given_str) : S_required;
    
    if (S_given_str !== "") {
      if (isNaN(S_given) || S_given <= 0) {
        showNotification("Given conductor area must be positive.", 'error');
        return;
      }
      if (S_given >= S_required) {
        html += '<p><strong style="color: #10b981;">Cable size is sufficient for the required area.</strong></p>';
        showNotification("Conductor calculation passed!", 'success');
      } else {
        html += '<p><strong style="color: #ef4444;">Cable undersized. Please choose the next available size.</strong></p>';
        showNotification("Cable undersized - check results!", 'warning');
      }
    }
    
    // Calculate I_AD for the given area
    const lnTerm = Math.log((theta_f + beta) / (theta_i + beta));
    const I_AD_A = K * S_given * Math.sqrt(lnTerm / t);
    const I_AD_calculated = I_AD_A / 1000;
    
    // Store data for PDF generation
    conductorData = {
      voltage: voltageKv,
      area: S_given,
      material: material,
      insulation: insulation,
      outer_sheath: outerSheath,
      scc_required: I_AD_kA,
      time: t,
      theta_i: theta_i,
      theta_f: theta_f,
      beta: beta,
      k_value: K,
      i_ad: I_AD_calculated.toFixed(3)
    };
    
  } else {
    const I_AD_kA_calc = calculateConductorCurrentFromArea();
    if (I_AD_kA_calc == null) return;
    html += `<p><strong>Adiabatic short-circuit current I<sub>AD</sub> for given area:</strong> <span style="font-size: 1.2rem; color: #2563eb;">${I_AD_kA_calc.toFixed(2)} kA</span></p>`;
    showNotification("Conductor calculation complete!", 'success');
    
    const S_given = parseFloat(S_given_str);
    
    // Store data for PDF generation
    conductorData = {
      voltage: voltageKv,
      area: S_given,
      material: material,
      insulation: insulation,
      outer_sheath: outerSheath,
      scc_required: I_AD_kA,
      time: t,
      theta_i: theta_i,
      theta_f: theta_f,
      beta: beta,
      k_value: K,
      i_ad: I_AD_kA_calc.toFixed(3)
    };
  }

  conductorCalculated = true;
  btnDownloadConductor.style.display = "inline-block";
  showResult(html);
});

// =============== STEP 3: SHEATH GEOMETRY ===============
function updateSheathGeometry() {
  const DoInput = document.getElementById("sheathOuterD");
  const DiInput = document.getElementById("sheathInnerD");
  const thicknessEl = document.getElementById("sheathThickness");
  const areaEl = document.getElementById("sheathAreaGiven");

  const Do = parseFloat(DoInput.value);
  const Di = parseFloat(DiInput.value);

  DoInput.style.borderColor = "";
  DiInput.style.borderColor = "";

  if (!isNaN(Do) && !isNaN(Di) && Do > 0 && Di > 0) {
    if (Do > Di) {
      const delta = (Do - Di) / 2;
      const area = (Math.PI / 4) * (Do * Do - Di * Di);
      thicknessEl.value = delta.toFixed(3);
      areaEl.value = area.toFixed(2);
    } else {
      thicknessEl.value = "";
      areaEl.value = "";
      DoInput.style.borderColor = "red";
      DiInput.style.borderColor = "red";
    }
  } else {
    thicknessEl.value = "";
    areaEl.value = "";
  }
}

document
  .getElementById("sheathOuterD")
  .addEventListener("input", updateSheathGeometry);
document
  .getElementById("sheathInnerD")
  .addEventListener("input", updateSheathGeometry);

// =============== SHEATH THERMAL HELPERS ===============
function getThermalConstants(materialType, materialName, voltageKv) {
  const group = THERMAL_CONSTANTS[materialType];
  if (!group) return null;
  const entry = group[materialName];
  if (!entry) return null;

  if (entry.rho !== undefined) return entry;

  if (materialType === "insulating") {
    if (materialName === "PVC" || materialName === "EPR") {
      return voltageKv <= 3 ? entry["<=3kV"] : entry[">3kV"];
    }
  } else if (materialType === "protective") {
    if (materialName === "PVC") {
      return voltageKv <= 35 ? entry["<=35kV"] : entry[">35kV"];
    }
  }
  return null;
}

function calculateM(
  insulationMaterial,
  outerSheathMaterial,
  sheathThickness,
  sheathMaterial,
  voltageKv,
  isOilFilled
) {
  const insulation = getThermalConstants(
    "insulating",
    insulationMaterial,
    voltageKv
  );
  const outerSheath = getThermalConstants(
    "protective",
    outerSheathMaterial,
    voltageKv
  );
  if (!insulation || !outerSheath) return null;

  const sigma2 = insulation.sigma;
  const rho2 = insulation.rho;
  const sigma3 = outerSheath.sigma;
  const rho3 = outerSheath.rho;

  const sheath = TABLE_I_SHEATHS[sheathMaterial];
  if (!sheath) return null;

  const sigma1 = sheath.sigmaC;
  const delta = sheathThickness;

  const F =
    isOilFilled === "yes"
      ? THERMAL_CONTACT_FACTOR["oil-filled"]
      : THERMAL_CONTACT_FACTOR.default;

  const sqrtTerm1 = Math.sqrt(sigma2 / rho2);
  const sqrtTerm2 = Math.sqrt(sigma3 / rho3);
  const numerator = sqrtTerm1 + sqrtTerm2;
  const denominator = 2 * sigma1 * delta * 1e-3;

  if (denominator === 0) return null;

  return (numerator / denominator) * F;
}

function calculateEpsilon(M, t) {
  if (M == null || isNaN(t) || t <= 0) return null;
  const MsqrtT = M * Math.sqrt(t);
  return (
    1 + 0.61 * MsqrtT - 0.069 * Math.pow(MsqrtT, 2) + 0.0043 * Math.pow(MsqrtT, 3)
  );
}

function calculateSheathAdiabaticArea(
  I_AD_kA,
  t,
  sheathMaterial,
  theta_i,
  theta_f
) {
  const mat = TABLE_I_SHEATHS[sheathMaterial];
  if (!mat) return null;

  const K = mat.K;
  const beta = mat.beta;

  const lnTerm = Math.log((theta_f + beta) / (theta_i + beta));
  if (lnTerm <= 0) return null;

  const I_AD_A = I_AD_kA * 1000;
  const s_sq = (I_AD_A ** 2 * t) / (K ** 2 * lnTerm);
  if (s_sq <= 0) return null;

  return Math.sqrt(s_sq);
}

// =============== SHEATH FORM SUBMIT ===============
sheathForm.addEventListener("submit", (e) => {
  e.preventDefault();
  console.log("=== SHEATH FORM SUBMITTED ===");
  // Don't hide result box - just update it

  const sheathMaterial = document
    .getElementById("sheathMaterial")
    .value.toLowerCase();
  console.log("Sheath material:", sheathMaterial);
  const voltageKv = parseFloat(
    document.getElementById("sheathVoltageKv").value
  );
  const I_AD_kA = parseFloat(
    document.getElementById("sheathSccKa").value
  );
  const t = parseFloat(
    document.getElementById("sheathTimeSec").value
  );
  const insulationMaterial =
    document.getElementById("insulationMaterial").value;
  const outerSheathMaterial =
    document.getElementById("outerSheathMaterial").value;
  const theta_i = parseFloat(
    document.getElementById("sheathThetaInitial").value
  );
  const theta_f = parseFloat(
    document.getElementById("sheathThetaFinal").value
  );
  const Do = parseFloat(document.getElementById("sheathOuterD").value);
  const Di = parseFloat(document.getElementById("sheathInnerD").value);
  const sheathThickness = parseFloat(
    document.getElementById("sheathThickness").value
  );
  const s_given = parseFloat(
    document.getElementById("sheathAreaGiven").value
  );
  
  console.log("=== SHEATH DIAMETER VALUES ===");
  console.log("Outer Diameter (Do):", Do);
  console.log("Inner Diameter (Di):", Di);
  console.log("Thickness:", sheathThickness);
  console.log("Given Area:", s_given);
  console.log("=== END DIAMETER VALUES ===");
  
  const conductorArea = parseFloat(document.getElementById("givenConductorArea").value) || 0;
  const conductorMaterial = document.getElementById("material").value;

  const isOilFilled = "no";

  console.log("Validation check:", {sheathMaterial, voltageKv, I_AD_kA, t, theta_i, theta_f, Do, Di, sheathThickness, s_given, insulationMaterial, outerSheathMaterial});
  
  if (
    !sheathMaterial ||
    [voltageKv, I_AD_kA, t, theta_i, theta_f, Do, Di, sheathThickness, s_given].some(
      (v) => isNaN(v)
    ) ||
    !insulationMaterial ||
    !outerSheathMaterial
  ) {
    console.log("Validation failed!");
    showNotification("Please fill all sheath inputs.", 'warning');
    return;
  }
  
  console.log("Validation passed, calculating...");

  if (!validateVoltageTime(voltageKv, t)) return;
  if (I_AD_kA <= 0) {
    showNotification("Short-circuit current must be positive.", 'error');
    return;
  }
  if (!(Do > Di && Do > 0 && Di > 0)) {
    showNotification("Outer diameter must be greater than inner diameter.", 'error');
    return;
  }

  const s_adiab = calculateSheathAdiabaticArea(
    I_AD_kA,
    t,
    sheathMaterial,
    theta_i,
    theta_f
  );
  if (s_adiab == null) {
    showNotification("Could not calculate adiabatic sheath area.", 'error');
    return;
  }

  const M = calculateM(
    insulationMaterial,
    outerSheathMaterial,
    sheathThickness,
    sheathMaterial,
    voltageKv,
    isOilFilled
  );
  if (M == null) {
    showNotification("Could not calculate M factor.", 'error');
    return;
  }

  const epsilon = calculateEpsilon(M, t);
  if (epsilon == null) {
    showNotification("Could not calculate ε factor.", 'error');
    return;
  }

  const s_required = s_adiab * epsilon;
  const i_non_ad = epsilon * s_adiab;

  let html = "<h6 style='color: #10b981;'>Sheath Calculation Results</h6><hr>";
  html += `<p><strong>Adiabatic area s<sub>adiab</sub>:</strong> ${s_adiab.toFixed(2)} mm²</p>`;
  html += `<p><strong>Non-adiabatic factor ε:</strong> ${epsilon.toFixed(3)}</p>`;
  html += `<p><strong>Required sheath area (non-adiabatic):</strong> <span style="font-size: 1.2rem; color: #2563eb;">${s_required.toFixed(2)} mm²</span></p>`;
  html += `<p><strong>Actual sheath area from D<sub>outer</sub>, D<sub>inner</sub>:</strong> ${s_given.toFixed(2)} mm²</p>`;

  if (s_given >= s_required) {
    html += '<p><strong style="color: #10b981;">Sheath size is sufficient for the required area.</strong></p>';
    showNotification("Sheath calculation passed!", 'success');
  } else {
    html += '<p><strong style="color: #ef4444;">Sheath undersized. Please choose the next available size.</strong></p>';
    showNotification("Sheath undersized - check results!", 'warning');
  }

  // Get thermal constants for PDF
  const insulation = getThermalConstants("insulating", insulationMaterial, voltageKv);
  const outerSheath = getThermalConstants("protective", outerSheathMaterial, voltageKv);
  const sheath = TABLE_I_SHEATHS[sheathMaterial];
  
  // Store data for PDF generation
  sheathData = {
    voltage: voltageKv,
    conductor_area: conductorArea,
    material: conductorMaterial,
    sheath_material: sheathMaterial.charAt(0).toUpperCase() + sheathMaterial.slice(1),
    insulation: insulationMaterial,
    outer_sheath: outerSheathMaterial,
    thickness: sheathThickness.toFixed(3),
    inner_d: Di.toFixed(2),
    outer_d: Do.toFixed(2),
    sheath_area: s_given.toFixed(2),
    scc_required: I_AD_kA,
    time: t,
    theta_i: theta_i,
    theta_f: theta_f,
    beta: sheath ? sheath.beta : 228,
    k_value: sheath ? sheath.K : 148,
    i_ad: s_adiab.toFixed(3),
    sigma1: sheath ? sheath.sigmaC : 2500000,
    sigma2: insulation ? insulation.sigma : 2400000,
    sigma3: outerSheath ? outerSheath.sigma : 2400000,
    rho2: insulation ? insulation.rho : 3.5,
    rho3: outerSheath ? outerSheath.rho : 3.5,
    f_factor: 0.7,
    m_factor: M.toFixed(3),
    epsilon: epsilon.toFixed(3),
    i_non_ad: i_non_ad.toFixed(3)
  };

  console.log("=== SHEATH DATA FOR PDF ===");
  console.log("inner_d:", sheathData.inner_d);
  console.log("outer_d:", sheathData.outer_d);
  console.log("thickness:", sheathData.thickness);
  console.log("sheath_area:", sheathData.sheath_area);
  console.log("=== END SHEATH DATA ===");

  console.log("Calculation complete, showing result");
  console.log("HTML to display:", html.substring(0, 100) + "...");
  sheathCalculated = true;
  btnDownloadSheath.style.display = "inline-block";
  console.log("About to call showResult");
  showResult(html);
  console.log("=== SHEATH CALCULATION COMPLETE ===");
});

// =============== PDF DOWNLOAD HANDLERS ===============

btnDownloadConductor.addEventListener("click", async () => {
  if (!conductorData) {
    showNotification("Please calculate conductor first.", 'warning');
    return;
  }
  
  try {
    const response = await fetch("/api/generate_conductor_pdf", {
      method: "POST",
      headers: {
        "Content-Type": "application/json"
      },
      body: JSON.stringify(conductorData)
    });
    
    if (!response.ok) {
      throw new Error("Failed to generate PDF");
    }
    
    const blob = await response.blob();
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = "Conductor_Calculation_Report.pdf";
    document.body.appendChild(a);
    a.click();
    window.URL.revokeObjectURL(url);
    document.body.removeChild(a);
    
    showNotification("Conductor report downloaded!", 'success');
  } catch (error) {
    console.error(error);
    showNotification("Failed to generate conductor PDF", 'error');
  }
});

btnDownloadSheath.addEventListener("click", async () => {
  if (!sheathData) {
    showNotification("Please calculate sheath first.", 'warning');
    return;
  }
  
  console.log("Sending sheath data:", sheathData);
  
  try {
    const response = await fetch("/api/generate_sheath_pdf", {
      method: "POST",
      headers: {
        "Content-Type": "application/json"
      },
      body: JSON.stringify(sheathData)
    });
    
    if (!response.ok) {
      throw new Error("Failed to generate PDF");
    }
    
    const blob = await response.blob();
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = "Sheath_Calculation_Report.pdf";
    document.body.appendChild(a);
    a.click();
    window.URL.revokeObjectURL(url);
    document.body.removeChild(a);
    
    showNotification("Sheath report downloaded!", 'success');
  } catch (error) {
    console.error(error);
    showNotification("Failed to generate sheath PDF", 'error');
  }
});

// Add CSS animations
const style = document.createElement('style');
style.textContent = `
  @keyframes slideIn {
    from {
      transform: translateX(100%);
      opacity: 0;
    }
    to {
      transform: translateX(0);
      opacity: 1;
    }
  }
  
  @keyframes slideOut {
    from {
      transform: translateX(0);
      opacity: 1;
    }
    to {
      transform: translateX(100%);
      opacity: 0;
    }
  }
`;
document.head.appendChild(style);
