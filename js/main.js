(function () {
  var config = window.BUH || { scriptUrl: "", target: 10000000000, mapUrl: "" };
  var money = new Intl.NumberFormat("en-US", {
    style: "currency",
    currency: "USD",
    maximumFractionDigits: 0
  });

  var amountEl = document.getElementById("amount-pledged");
  var supportersEl = document.getElementById("supporters");
  var remainingEl = document.getElementById("remaining");
  var countriesEl = document.getElementById("countries");
  var recentEl = document.getElementById("recent");
  var largestEl = document.getElementById("largest");
  var countrySelect = document.getElementById("country");
  var form = document.getElementById("pledge-form");
  var formError = document.getElementById("form-error");
  var formFields = document.getElementById("form-fields");
  var success = document.getElementById("success");
  var share = document.getElementById("share-line");
  var panel = document.getElementById("panel");
  var overlay = document.getElementById("overlay");
  var displayName = document.getElementById("display-name");
  var isPublic = document.getElementById("is-public");

  var map;
  var geojson;
  var countryLayer;
  var latestStats = emptyStats();

  document.getElementById("open-pledge").addEventListener("click", openPanel);
  document.getElementById("close-pledge").addEventListener("click", closePanel);
  overlay.addEventListener("click", closePanel);
  document.addEventListener("keydown", function (event) {
    if (event.key === "Escape") closePanel();
  });
  isPublic.addEventListener("change", toggleDisplayName);
  form.addEventListener("submit", onSubmit);
  document.getElementById("copy-share").addEventListener("click", function () {
    if (navigator.clipboard && navigator.clipboard.writeText) {
      navigator.clipboard.writeText(share.value);
      return;
    }
    share.select();
    document.execCommand("copy");
  });

  render(latestStats);
  loadMap();
  loadStats();

  function emptyStats() {
    return { pledged: 0, supporters: 0, countries: {}, recent: [], largest: [] };
  }

  function render(stats) {
    latestStats = stats || emptyStats();
    var pledged = Number(latestStats.pledged) || 0;
    var supporters = Number(latestStats.supporters) || 0;
    var countryMap = latestStats.countries || {};
    var countryCount = Object.keys(countryMap).filter(function (name) {
      return countryMap[name] > 0;
    }).length;
    amountEl.textContent = money.format(pledged);
    supportersEl.textContent = String(supporters);
    remainingEl.textContent = money.format(Math.max(0, config.target - pledged));
    countriesEl.textContent = String(countryCount);
    renderList(recentEl, latestStats.recent);
    renderList(largestEl, latestStats.largest);
    paintMap();
  }

  function renderList(list, rows) {
    list.innerHTML = "";
    if (!rows || !rows.length) {
      var empty = document.createElement("li");
      empty.className = "empty";
      empty.textContent = "No public pledges yet.";
      list.appendChild(empty);
      return;
    }
    rows.forEach(function (row) {
      var item = document.createElement("li");
      var who = document.createElement("span");
      who.className = "who";
      who.textContent = row.name;
      var sum = document.createElement("span");
      sum.className = "sum";
      sum.textContent = money.format(Number(row.amount) || 0);
      var meta = document.createElement("span");
      meta.className = "meta";
      meta.textContent = row.country || "";
      item.appendChild(who);
      item.appendChild(sum);
      item.appendChild(meta);
      list.appendChild(item);
    });
  }

  function loadStats() {
    if (!config.scriptUrl) return;
    jsonp(config.scriptUrl, { action: "stats" }).then(function (data) {
      if (data && data.stats) render(data.stats);
    }).catch(function () {});
  }

  function loadMap() {
    if (typeof L === "undefined") {
      countrySelect.innerHTML = '<option value="">Country list unavailable</option>';
      return;
    }
    map = L.map("map", {
      scrollWheelZoom: false,
      zoomControl: false,
      attributionControl: true
    }).setView([20, 10], 2);
    L.control.zoom({ position: "bottomright" }).addTo(map);
    fetch(config.mapUrl)
      .then(function (response) { return response.json(); })
      .then(function (data) {
        geojson = data;
        fillCountries(data);
        paintMap();
      })
      .catch(function () {
        countrySelect.innerHTML = '<option value="United Kingdom">United Kingdom</option>';
      });
  }

  function fillCountries(data) {
    var names = [];
    (data.features || []).forEach(function (feature) {
      var name = countryName(feature);
      if (name && names.indexOf(name) === -1) names.push(name);
    });
    names.sort();
    countrySelect.innerHTML = '<option value="">Country</option>';
    names.forEach(function (name) {
      var option = document.createElement("option");
      option.value = name;
      option.textContent = name;
      countrySelect.appendChild(option);
    });
  }

  function paintMap() {
    if (!map || !geojson) return;
    if (countryLayer) map.removeLayer(countryLayer);
    var counts = latestStats.countries || {};
    var max = 1;
    Object.keys(counts).forEach(function (name) {
      if (counts[name] > max) max = counts[name];
    });
    countryLayer = L.geoJSON(geojson, {
      style: function (feature) {
        var count = counts[countryName(feature)] || 0;
        return {
          color: "#f4f0e8",
          weight: 0.6,
          fillColor: shade(count, max),
          fillOpacity: 1
        };
      },
      onEachFeature: function (feature, layer) {
        var name = countryName(feature);
        var count = counts[name] || 0;
        layer.bindTooltip(name + (count ? " · " + count : ""), { sticky: true });
      }
    }).addTo(map);
  }

  function countryName(feature) {
    var props = feature.properties || {};
    return props.NAME || props.ADMIN || props.name || "";
  }

  function shade(count, max) {
    if (!count) return "#e4ddd2";
    var t = Math.max(0.18, Math.min(1, count / max));
    var r = Math.round(232 - t * 143);
    var g = Math.round(214 - t * 185);
    var b = Math.round(206 - t * 177);
    return "rgb(" + r + "," + g + "," + b + ")";
  }

  function openPanel() {
    overlay.hidden = false;
    panel.hidden = false;
    document.body.style.overflow = "hidden";
    document.getElementById("first-name").focus();
  }

  function closePanel() {
    overlay.hidden = true;
    panel.hidden = true;
    document.body.style.overflow = "";
  }

  function toggleDisplayName() {
    displayName.required = isPublic.checked;
  }

  function onSubmit(event) {
    event.preventDefault();
    formError.textContent = "";
    if (!config.scriptUrl) {
      formError.textContent = "The pledge book is not connected yet.";
      return;
    }
    var payload = {
      firstName: value("first-name"),
      lastName: value("last-name"),
      age: value("age"),
      email: value("email"),
      country: value("country"),
      amount: value("amount"),
      payment: value("payment"),
      isPublic: isPublic.checked,
      agreed: document.getElementById("agreed").checked,
      displayName: value("display-name"),
      company: value("company")
    };
    var button = document.getElementById("submit-pledge");
    button.disabled = true;
    postPledge(payload).then(function (data) {
      button.disabled = false;
      if (!data || !data.ok) {
        formError.textContent = (data && data.error) || "Could not save the pledge.";
        return;
      }
      if (data.stats) render(data.stats);
      else loadStats();
      share.value = "I pledged to Bring United Home. No money has been paid.";
      formFields.hidden = true;
      success.hidden = false;
    }).catch(function () {
      button.disabled = false;
      formError.textContent = "Could not reach the pledge book.";
    });
  }

  function postPledge(payload) {
    return fetch(config.scriptUrl, {
      method: "POST",
      redirect: "follow",
      headers: { "Content-Type": "text/plain;charset=utf-8" },
      body: JSON.stringify(payload)
    }).then(function (response) { return response.json(); });
  }

  function jsonp(url, params) {
    return new Promise(function (resolve, reject) {
      var callback = "buh_" + Date.now() + "_" + Math.floor(Math.random() * 1000);
      var script = document.createElement("script");
      var timer = setTimeout(function () {
        cleanup();
        reject(new Error("timeout"));
      }, 12000);
      window[callback] = function (data) {
        cleanup();
        resolve(data);
      };
      function cleanup() {
        clearTimeout(timer);
        delete window[callback];
        if (script.parentNode) script.parentNode.removeChild(script);
      }
      var query = Object.keys(params).map(function (key) {
        return encodeURIComponent(key) + "=" + encodeURIComponent(params[key]);
      }).join("&");
      var join = url.indexOf("?") === -1 ? "?" : "&";
      script.src = url + join + query + "&callback=" + callback;
      script.onerror = function () {
        cleanup();
        reject(new Error("network"));
      };
      document.body.appendChild(script);
    });
  }

  function value(id) {
    return document.getElementById(id).value.trim();
  }
})();
