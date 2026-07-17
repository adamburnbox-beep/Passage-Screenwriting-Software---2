# Passage web app — build from the repository root:
#   docker build -t passage-web .
FROM mcr.microsoft.com/dotnet/sdk:9.0 AS build
WORKDIR /src
COPY Directory.Build.props ./
COPY Passage/Passage.Core/ Passage/Passage.Core/
COPY Passage/Passage.Parser/ Passage/Passage.Parser/
COPY Passage/Passage.Export/ Passage/Passage.Export/
COPY Passage/Passage.Web/ Passage/Passage.Web/
RUN dotnet publish Passage/Passage.Web/Passage.Web.csproj -c Release -o /app/publish

FROM mcr.microsoft.com/dotnet/aspnet:9.0 AS runtime
WORKDIR /app
COPY --from=build /app/publish .

# Scripts are stored on this volume; mount it to keep them across updates.
ENV ASPNETCORE_URLS=http://+:8080 \
    PASSAGE_DATA_DIR=/data
VOLUME ["/data"]
EXPOSE 8080

ENTRYPOINT ["dotnet", "Passage.Web.dll"]
