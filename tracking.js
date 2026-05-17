(function () {
  const ORDER_API_URL = "https://script.google.com/macros/s/AKfycbyPC-bIh5AYPp0PCTuQoC-XEMmmwJLVAHN7cBLEa8euDTSDgUtKGqGQQpth1bT_QtKS/exec";
  const HEARTBEAT_MS = 60000;

  function buildPayload(user, extras) {
    const details = extras || {};
    return {
      uid: user.uid || "",
      email: user.email || "",
      displayName: user.displayName || "",
      page: details.page || (location.pathname.split("/").pop() || "index.html"),
      pageLabel: details.pageLabel || document.title,
      note: details.note || "",
      eventType: details.eventType || "page_view",
      lastAction: details.lastAction || "active"
    };
  }

  async function postJson(body) {
    try {
      await fetch(ORDER_API_URL, {
        method: "POST",
        body: JSON.stringify(body)
      });
    } catch (error) {
      console.error("Tracking request failed", error);
    }
  }

  function startUserTracking(user, options) {
    if (!user || !user.uid) {
      return function () {};
    }

    const pagePayload = buildPayload(user, options || {});
    postJson(Object.assign({ action: "trackVisit" }, pagePayload));
    postJson(Object.assign({ action: "trackPresence" }, pagePayload));

    const heartbeat = setInterval(function () {
      postJson(Object.assign({
        action: "trackPresence"
      }, buildPayload(user, {
        page: pagePayload.page,
        pageLabel: pagePayload.pageLabel,
        lastAction: "heartbeat"
      })));
    }, HEARTBEAT_MS);

    const handleVisibility = function () {
      if (!document.hidden) {
        postJson(Object.assign({
          action: "trackPresence"
        }, buildPayload(user, {
          page: pagePayload.page,
          pageLabel: pagePayload.pageLabel,
          lastAction: "visible"
        })));
      }
    };

    document.addEventListener("visibilitychange", handleVisibility);

    return function () {
      clearInterval(heartbeat);
      document.removeEventListener("visibilitychange", handleVisibility);
    };
  }

  function trackLoginSuccess(user, note) {
    if (!user || !user.uid) {
      return;
    }

    postJson(Object.assign({
      action: "trackVisit"
    }, buildPayload(user, {
      page: "login.html",
      pageLabel: "หน้าเข้าสู่ระบบ",
      eventType: "login_success",
      note: note || ""
    })));
  }

  function trackLogout(user, page, pageLabel) {
    if (!user || !user.uid) {
      return;
    }

    postJson(Object.assign({
      action: "trackVisit"
    }, buildPayload(user, {
      page: page || (location.pathname.split("/").pop() || "index.html"),
      pageLabel: pageLabel || document.title,
      eventType: "logout",
      lastAction: "logout"
    })));
  }

  window.KopcTracking = {
    startUserTracking: startUserTracking,
    trackLoginSuccess: trackLoginSuccess,
    trackLogout: trackLogout
  };
})();


