workflow "Deploy" {
  on = "push"
  resolves = ["Deploy to PSGallery"]
}


action "Deploy to PSGallery" {
  uses = "./.github/gallerydeploy"
  secrets = ["PSGALLERY_DEPLOY_KEY"]
  env = {
    
  }
}