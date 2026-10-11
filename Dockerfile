# OfficeHere web sürümü: derle, nginx ile statik yayınla.
FROM node:22-alpine AS derleme
RUN corepack enable
WORKDIR /src
COPY . .
ENV ELECTRON_SKIP_BINARY_DOWNLOAD=1
RUN pnpm install --frozen-lockfile && pnpm web

FROM nginx:1.27-alpine
COPY --from=derleme /src/dist-web /usr/share/nginx/html
COPY docker/nginx.conf /etc/nginx/conf.d/default.conf
EXPOSE 80
