require 'zoho_mailer'

module ActionMailer
  module DeliveryMethods
    class Zoho
      def initialize(settings)
        @settings = settings || {}
        configure_zoho_mailer
      end

      def deliver!(mail)
        # Convert ActionMailer::MessageDelivery to ZohoMailer format
        to_addresses = Array(mail.to).join(',')
        cc_addresses = Array(mail.cc).join(',') if mail.cc.present?
        bcc_addresses = Array(mail.bcc).join(',') if mail.bcc.present?

        # Extract from address
        from_address = extract_from_address(mail)

        # Get email body (prefer HTML, fallback to text)
        body = extract_body(mail)

        # Prepare attachments
        attachments = prepare_attachments(mail)

        params = {
          from: from_address,
          to: to_addresses,
          cc: cc_addresses,
          bcc: bcc_addresses,
          subject: mail.subject,
          body: body,
          attachments: attachments
        }.compact

        result = ZohoMailer.send_email(params)

        unless result
          raise StandardError, "Failed to send email via Zoho Mailer"
        end

        result
      end

      private

      def configure_zoho_mailer
        ZohoMailer.setup do |config|
          config.client_id = @settings[:client_id] || ENV['ZOHO_CLIENT_ID']
          config.client_secret = @settings[:client_secret] || ENV['ZOHO_CLIENT_SECRET']
          config.refresh_token = @settings[:refresh_token] || ENV['ZOHO_REFRESH_TOKEN']
          config.from_email = @settings[:from_email] || ENV['ZOHO_FROM_EMAIL'] || ENV['MAILER_SENDER_EMAIL']
          config.token_url = @settings[:token_url] || ENV['ZOHO_TOKEN_URL'] || 'https://accounts.zoho.com/oauth/v2/token'
          config.mail_api_url = @settings[:mail_api_url] || ENV['ZOHO_MAIL_API_URL'] || 'https://mail.zoho.com/api/accounts'
          config.token_store = (@settings[:token_store] || ENV['ZOHO_TOKEN_STORE'] || 'file').to_sym
          config.token_file_path = @settings[:token_file_path] || ENV['ZOHO_TOKEN_FILE_PATH'] || Rails.root.join('tmp', 'zoho_token.json').to_s
          config.timeout = (@settings[:timeout] || ENV['ZOHO_TIMEOUT'] || 30).to_i
        end
      end

      def extract_from_address(mail)
        from = mail.from
        return from.first if from.is_a?(Array) && from.any?

        # Try to extract from header
        from_header = mail.header['from']
        return from_header.value if from_header

        # Fallback to default
        ENV['ZOHO_FROM_EMAIL'] || ENV['MAILER_SENDER_EMAIL'] || 'noreply@chatwoot.com'
      end

      def extract_body(mail)
        # Prefer HTML part
        if mail.html_part
          mail.html_part.body.decoded
        elsif mail.text_part
          mail.text_part.body.decoded
        elsif mail.body
          mail.body.decoded
        else
          ''
        end
      end

      def prepare_attachments(mail)
        return [] unless mail.attachments.any?

        attachments = []
        mail.attachments.each do |attachment|
          # Save attachment temporarily and upload to Zoho
          temp_file = save_attachment_temp(attachment)
          next unless temp_file

          begin
            # Upload attachment to Zoho and get attachment info
            attachment_info = ZohoMailer.send(:upload_attachment, temp_file)
            attachments << attachment_info if attachment_info
          ensure
            # Clean up temp file
            File.delete(temp_file) if File.exist?(temp_file)
          end
        end
        attachments
      end

      def save_attachment_temp(attachment)
        return nil unless attachment

        temp_dir = Rails.root.join('tmp', 'zoho_attachments')
        FileUtils.mkdir_p(temp_dir)

        filename = attachment.filename || "attachment_#{SecureRandom.hex(8)}"
        temp_file = File.join(temp_dir, filename)
        File.binwrite(temp_file, attachment.body.decoded)

        temp_file
      rescue StandardError => e
        Rails.logger.error "Failed to save attachment for Zoho: #{e.message}"
        nil
      end
    end
  end
end

