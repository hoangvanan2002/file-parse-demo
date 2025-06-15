FROM eclipse-temurin:21-jdk-alpine

WORKDIR /app

# Copy JAR file
COPY target/file-parse-demo-0.0.1-SNAPSHOT.jar app.jar

# Expose Spring Boot port
EXPOSE 8080

# Run the application
ENTRYPOINT ["java", "-jar", "app.jar"]
