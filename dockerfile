# syntax=docker/dockerfile:1
FROM mcr.microsoft.com/dotnet/sdk:6.0 AS build-env
WORKDIR /app
ARG APP_VERSION


# Copy everything else and build
COPY . ./
RUN dotnet restore

# Run tests
RUN dotnet test JsonToWord.Tests/JsonToWord.Tests.csproj

RUN if [ -n "$APP_VERSION" ]; then \
      dotnet publish -c Release -o out -p:Version="$APP_VERSION" -p:InformationalVersion="$APP_VERSION"; \
    else \
      dotnet publish -c Release -o out; \
    fi

# Build runtime image
FROM mcr.microsoft.com/dotnet/sdk:6.0
ENV ASPNETCORE_URLS=http://+:5000  
WORKDIR /app
COPY --from=build-env /app/out .
ENTRYPOINT ["dotnet", "JsonToWord.dll"]
