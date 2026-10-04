// Fills links that point at the GitHub repository, derived from the Pages address
// (https://<user>.github.io/<repo>/ maps to https://github.com/<user>/<repo>).
(function () {
  var host = location.hostname.split(".")[0];
  var repo = location.pathname.split("/").filter(Boolean)[0] || "";
  var url = location.hostname.endsWith("github.io") ? "https://github.com/" + host + "/" + repo : "#";
  document.querySelectorAll("[data-repo-link]").forEach(function (a) {
    a.href = url + (a.getAttribute("data-repo-link") || "");
  });
})();
