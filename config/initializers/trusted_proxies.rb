# frozen_string_literal: true

# Configuración de proxies confiables para que Rails pueda leer correctamente
# las headers X-Forwarded-For y X-Real-IP cuando hay un proxy (nginx) delante
#
# Este initializer es necesario cuando Chatwoot está detrás de un proxy reverso
# (nginx, Apache, load balancer, etc.) para que rack-attack y otras funcionalidades
# puedan detectar correctamente la IP real del cliente en lugar de la IP del proxy.
#
# Sin esta configuración, todas las peticiones aparecerían como si vinieran de
# 127.0.0.1 o la IP del proxy, causando que rack-attack bloquee todas las
# peticiones con error 429 (Too Many Requests).
Rails.application.config.action_dispatch.trusted_proxies = if Rails.env.production?
  # En producción, confiar en localhost y redes privadas
  # Esto permite que nginx pase las IPs reales a través de X-Forwarded-For
  [
    IPAddr.new('127.0.0.1'),
    IPAddr.new('::1'),
    IPAddr.new('10.0.0.0/8'),
    IPAddr.new('172.16.0.0/12'),
    IPAddr.new('192.168.0.0/16'),
  ]
else
  # En desarrollo y test, usar la configuración por defecto de Rails
  ActionDispatch::RemoteIp::TRUSTED_PROXIES
end

