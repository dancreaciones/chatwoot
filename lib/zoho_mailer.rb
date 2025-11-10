require 'net/http'
require 'uri'
require 'json'
require 'active_support/core_ext/hash'

module ZohoMailer
  class Error < StandardError; end

  class ConfigurationError < Error; end

  class APIError < Error; end

  class Configuration
    attr_accessor :client_id, :client_secret, :refresh_token, :from_email,
                  :token_url, :mail_api_url, :token_store, :token_file_path,
                  :timeout

    def initialize
      @token_url = 'https://accounts.zoho.com/oauth/v2/token'
      @mail_api_url = 'https://mail.zoho.com/api/accounts'
      @token_store = :file
      @timeout = 30
      @access_token = nil
      @access_token_expiration = nil
    end
  end

  class << self
    def config
      @config ||= Configuration.new
    end

    def setup
      yield(config) if block_given?
      validate_configuration
    end

    def send_email(params)
      validate_email_params(params)
      email_data = build_email_data(params)

      with_error_handling do
        account_id = fetch_account_id
        access_token = get_access_token
        response = make_email_request(account_id, access_token, email_data)
        handle_email_response(response)
      end
    end

    def send_email_with_file(to:, cc:, subject:, body:, file_path:)
      attachment_info = upload_attachment(file_path)
      send_email(
        to: to,
        cc: cc,
        subject: subject,
        body: body,
        attachments: [attachment_info]
      )
    end

    private

    def validate_configuration
      required_keys = %i[client_id client_secret refresh_token]
      missing_keys = required_keys.select { |k| config.send(k).to_s.empty? }

      if missing_keys.any?
        raise ConfigurationError, "Missing configuration: #{missing_keys.join(', ')}"
      end

      if config.token_store == :file && config.token_file_path.to_s.empty?
        raise ConfigurationError, "Token file path is required for :file token store"
      end
    end

    def validate_email_params(params)
      required_params = %i[to subject body]
      missing_params = required_params - params.keys.map(&:to_sym)

      if missing_params.any?
        raise ConfigurationError, "Missing required parameters: #{missing_params.join(', ')}"
      end
    end

    def build_email_data(params)
      {
        fromAddress: params[:from] || config.from_email,
        toAddress: params[:to],
        ccAddress: params[:cc],
        bccAddress: params[:bcc],
        subject: params[:subject],
        content: params[:body],
        askReceipt: params[:receipt] ? "yes" : "no",
        attachments: prepare_attachments(params[:attachments])
      }.compact
    end

    def prepare_attachments(attachments)
      return unless attachments.is_a?(Array)

      attachments.map do |attachment|
        if attachment.is_a?(String)
          { filepath: attachment, filename: File.basename(attachment) }
        else
          attachment
        end
      end.compact
    end

    def make_email_request(account_id, access_token, email_data)
      uri = URI.parse("#{config.mail_api_url}/#{account_id}/messages")
      Rails.logger.info("ZohoMailer: Email data: #{email_data.to_json}")

      http = Net::HTTP.new(uri.host, uri.port)
      http.use_ssl = true
      http.verify_mode = OpenSSL::SSL::VERIFY_NONE
      http.read_timeout = config.timeout

      request = Net::HTTP::Post.new(uri)
      request['Content-Type'] = 'application/json'
      request.body = email_data.to_json
      request['Authorization'] = "Zoho-oauthtoken #{access_token}"

      response = http.request(request)
      Rails.logger.info("ZohoMailer: Response: #{response}")

      case response
      when Net::HTTPSuccess
        Rails.logger.info("ZohoMailer: Email enviado correctamente")
        true
      else
        raise APIError, "Falló el envío del email: #{response.code} - #{response.body}"
      end
    end

    def upload_attachment(file_path)
      account_id = fetch_account_id
      access_token = get_access_token

      uri = URI.parse("https://mail.zoho.com/api/accounts/#{account_id}/messages/attachments?uploadType=multipart")

      require 'net/http/post/multipart'
      require 'stringio'

      # Detect MIME type
      mime_type = detect_mime_type(file_path)

      file_io = UploadIO.new(File.open(file_path), mime_type, File.basename(file_path))
      request = Net::HTTP::Post::Multipart.new(uri.path + "?uploadType=multipart", { 'attach' => file_io })
      request['Authorization'] = "Zoho-oauthtoken #{access_token}"

      http = Net::HTTP.new(uri.host, uri.port)
      http.use_ssl = true
      http.verify_mode = OpenSSL::SSL::VERIFY_NONE

      response = http.request(request)

      unless response.is_a?(Net::HTTPSuccess)
        raise APIError, "Falló la carga del adjunto: #{response.code} - #{response.body}"
      end

      data = JSON.parse(response.body)['data'].first
      {
        storeName: data['storeName'],
        attachmentPath: data['attachmentPath'],
        attachmentName: data['attachmentName']
      }
    end

    def detect_mime_type(file_path)
      require 'mime/types'
      mime_type = MIME::Types.type_for(file_path).first
      mime_type ? mime_type.content_type : 'application/octet-stream'
    rescue LoadError
      # Fallback if mime-types gem is not available
      ext = File.extname(file_path).downcase
      case ext
      when '.pdf'
        'application/pdf'
      when '.jpg', '.jpeg'
        'image/jpeg'
      when '.png'
        'image/png'
      when '.gif'
        'image/gif'
      when '.doc'
        'application/msword'
      when '.docx'
        'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
      when '.xls'
        'application/vnd.ms-excel'
      when '.xlsx'
        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
      else
        'application/octet-stream'
      end
    end

    def handle_email_response(response)
      case response
      when Net::HTTPSuccess
        Rails.logger.info("ZohoMailer: Email sent successfully")
        true
      else
        raise APIError, "Failed to send email: #{response.code} - #{response.body}"
      end
    end

    def fetch_account_id
      access_token = get_access_token
      uri = URI.parse(config.mail_api_url)

      http = Net::HTTP.new(uri.host, uri.port)
      http.use_ssl = true
      http.verify_mode = OpenSSL::SSL::VERIFY_NONE
      http.read_timeout = config.timeout

      request = Net::HTTP::Get.new(uri)
      request['Authorization'] = "Zoho-oauthtoken #{access_token}"

      response = http.request(request)

      unless response.is_a?(Net::HTTPSuccess)
        raise APIError, "Failed to fetch account ID: #{response.body}"
      end

      data = JSON.parse(response.body)['data']
      raise APIError, "No accounts data in response" unless data.is_a?(Array) && data.any?

      account_id = data.first['accountId']
      raise APIError, "No accountId in response" unless account_id

      account_id
    end

    def get_access_token
      # Verifica si el token actual es válido
      if @access_token && @access_token_expiration && Time.now < @access_token_expiration
        return @access_token
      end

      refresh_token = load_refresh_token
      params = {
        refresh_token: refresh_token,
        client_id: config.client_id,
        client_secret: config.client_secret,
        grant_type: 'refresh_token'
      }

      uri = URI.parse(config.token_url)
      response = Net::HTTP.post_form(uri, params)

      unless response.is_a?(Net::HTTPSuccess)
        raise APIError, "Failed to get access token: #{response.body}"
      end

      data = JSON.parse(response.body)
      Rails.logger.info("ZohoMailer: Response body: #{response}")
      Rails.logger.info("ZohoMailer: Response data: #{data}")

      raise APIError, "No access token in response" unless data['access_token']

      # Guarda el token de acceso y establece su tiempo de expiración (30 minutos)
      @access_token = data['access_token']
      @access_token_expiration = Time.now + 30 * 60

      save_refresh_token(data['refresh_token']) if data['refresh_token']

      @access_token
    end

    def load_refresh_token
      case config.token_store
      when :file
        load_refresh_token_from_file
      when :memory
        config.refresh_token
      else
        raise ConfigurationError, "Invalid token store: #{config.token_store}"
      end
    end

    def load_refresh_token_from_file
      return config.refresh_token unless File.exist?(config.token_file_path)

      data = JSON.parse(File.read(config.token_file_path))
      data['refresh_token'] || config.refresh_token
    rescue JSON::ParserError => e
      Rails.logger.error "Error parsing refresh token file: #{e.message}"
      config.refresh_token
    end

    def save_refresh_token(token)
      case config.token_store
      when :file
        save_refresh_token_to_file(token)
      when :memory
        config.refresh_token = token
      end
    end

    def save_refresh_token_to_file(token)
      FileUtils.mkdir_p(File.dirname(config.token_file_path))
      File.write(config.token_file_path, { refresh_token: token }.to_json)
    end

    def with_error_handling
      yield
    rescue ConfigurationError => e
      Rails.logger.error "ZohoMailer Configuration Error: #{e.message}"
      false
    rescue APIError => e
      Rails.logger.error "ZohoMailer API Error: #{e.message}"
      false
    rescue StandardError => e
      Rails.logger.error "ZohoMailer Unexpected Error: #{e.message}"
      false
    end
  end
end

