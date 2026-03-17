import { Application } from "@hotwired/stimulus"
import DashboardController from "./controllers/dashboard_controller"
import SettingsController from "./controllers/settings_controller"
import XlsxUploadController from "./controllers/xlsx_upload_controller"

const application = Application.start()
application.register("dashboard", DashboardController)
application.register("settings", SettingsController)
application.register("xlsx-upload", XlsxUploadController)
