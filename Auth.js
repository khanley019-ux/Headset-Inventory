function getPage(page) {

  switch (page) {

    case "dashboard":
      return HtmlService.createHtmlOutputFromFile("PAGE_Dashboard").getContent();

    default:
      return `
        <div class="container py-5">

            <h2>404</h2>

            <p>Page not found.</p>

        </div>
      `;

  }

}