# 📊 DataWeaver - Timesheet Processor

**DataWeaver** is a Java + Spring Boot application that simplifies how managers process team timesheets. By accepting bulk Excel files as input, it cleanly organizes data and returns personalized timesheets—per member, per day—with a summary sheet for quick insights.

---

## 🚀 Features

- 📥 Accepts timesheet data as a single Excel file via API
- 📅 Splits entries by individual and day-wise from 1st to last of the month
- 📤 Returns:
  - One Excel file per team member
  - A summary Excel sheet with aggregate data
- 🔁 Handles both upload and download of `.xlsx` format
- 🧹 Applies business logic for cleaner separation and reporting

---

## 🧰 Tech Stack

| Technology              | Purpose                                      |
|------------------------|----------------------------------------------|
| **Java 17**            | Core language                                |
| **Spring Boot 3.4.4**  | Backend framework                            |
| **Spring Web**         | RESTful API development                      |
| **Apache POI 5.2.2**   | Excel (.xlsx) read/write operations          |
| **Springdoc OpenAPI**  | Swagger-based API documentation              |
| **Lombok**             | Boilerplate-free Java annotations            |
| **Spring Boot Actuator** | App health metrics & monitoring            |
| **Maven**              | Build and dependency management              |

---

## 📁 API Overview

### `POST /data-weaver/generate-excel`

Upload a team Excel timesheet to receive processed outputs.

- **Request:** Multipart file (`.xlsx`)
- **Response:** Multipart file (`.xlsx`):
  - One Excel per employee
  - One summary sheet

 ### `GET /data-weaver/health-check`

end point to check the health of the application

- **Request:** No request body
- **Response:** a string 

### Swagger UI

View and interact with APIs:

