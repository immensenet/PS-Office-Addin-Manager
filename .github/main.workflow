workflow "Deploy to PSGallery" {
  on = "push"
  resolves = ["Deploy to PSGallery"]
}


action "Deploy to Azure" {
  uses = "./.github/gallerydeploy"
  secrets = ["PSGALLERY_DEPLOY_KEY"]
  env = {
    
  }
}