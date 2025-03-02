# syntax=docker/dockerfile:1
FROM mcr.microsoft.com/dotnet/sdk:6.0 AS build-env
WORKDIR /app


# Copy everything else and build
COPY . ./
RUN dotnet restore

# Run tests
RUN dotnet test JsonToWord.Tests/JsonToWord.Tests.csproj

RUN dotnet publish -c Release -o out

# Build runtime image
FROM mcr.microsoft.com/dotnet/sdk:6.0
ENV ASPNETCORE_URLS=http://+:5000  
WORKDIR /app
COPY --from=build-env /app/out .
ENTRYPOINT ["dotnet", "JsonToWord.dll"]