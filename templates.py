# -*- coding: utf-8 -*-
"""
Corpo HTML do e-mail enviado ao fornecedor, tema azul.

Suporta dois idiomas: inglês (padrão, para fornecedores internacionais) e
português do Brasil (para fornecedores nacionais). O idioma é escolhido pelo
chamador com base no "Country Code" do fornecedor no Excel do SRM - "BR" usa
o texto em português, qualquer outro valor usa o texto em inglês.
"""
from reminder_engine import EXPIRED, NO_DATE


def _body_en(supplier_name, cert_type_label, expiry_date_str, days_left, tier, company_name):
    if tier == EXPIRED:
        title = "Certificate Expired"
        accent_color = "#c0392b"
        status_message = (
            f"The certificate has been <strong>expired for {abs(days_left)} day(s)</strong>. "
            f"Please regularize this as soon as possible."
        )
    else:
        title = "Certificate Renewal"
        accent_color = "#1f6fb2"
        status_message = f"The certificate expires in <strong>{days_left} day(s)</strong>."

    return f"""\
<html>
<body style="font-family: 'Segoe UI', Arial, sans-serif; background-color:#eef4fb; margin:0; padding:0;">
  <table width="100%" cellpadding="0" cellspacing="0" style="background-color:#eef4fb; padding:30px 0;">
    <tr><td align="center">
      <table width="600" cellpadding="0" cellspacing="0"
             style="background-color:#ffffff; border-radius:8px; overflow:hidden; box-shadow:0 2px 8px rgba(13,59,102,0.15);">
        <tr>
          <td style="background-color:#0d3b66; padding:24px 32px;">
            <span style="color:#ffffff; font-size:20px; font-weight:600;">{company_name}</span><br>
            <span style="color:#a9c6e8; font-size:13px;">Supplier Quality Management</span>
          </td>
        </tr>
        <tr>
          <td style="padding:32px;">
            <h2 style="color:{accent_color}; margin-top:0; font-family: 'Segoe UI', Arial, sans-serif;">{title}</h2>
            <p style="color:#333333; font-size:15px; line-height:1.5;">
              Dear <strong>{supplier_name}</strong>,
            </p>
            <p style="color:#333333; font-size:15px; line-height:1.5;">
              We have identified that this supplier's <strong>{cert_type_label}</strong> certificate
              is valid until <strong>{expiry_date_str}</strong>.
            </p>
            <p style="color:#333333; font-size:15px; line-height:1.5;">{status_message}</p>
            <div style="background-color:#f0f6fc; border-left:4px solid {accent_color}; padding:14px 18px; margin:20px 0; border-radius:4px;">
              <p style="margin:0; color:#0d3b66; font-size:14px; line-height:1.5;">
                Please send us an updated copy of the renewed certificate as soon as possible,
                so we can keep your approved-supplier record up to date.
              </p>
            </div>
            <p style="color:#333333; font-size:15px; line-height:1.5;">
              If you have any questions, please contact our Supplier Quality team.
            </p>
            <p style="color:#333333; font-size:15px; line-height:1.5;">
              Kind regards,<br>Supplier Quality Team
            </p>
          </td>
        </tr>
        <tr>
          <td style="background-color:#0d3b66; padding:14px 32px;">
            <span style="color:#a9c6e8; font-size:11px;">
              This is an automated message from the supplier certificate management system.
            </span>
          </td>
        </tr>
      </table>
    </td></tr>
  </table>
</body>
</html>
"""


def _body_missing_pt(supplier_name, cert_type_label, company_name):
    """E-mail enviado quando não há data de validade cadastrada: pergunta se o
    fornecedor possui o certificado."""
    return f"""\
<html>
<body style="font-family: 'Segoe UI', Arial, sans-serif; background-color:#eef4fb; margin:0; padding:0;">
  <table width="100%" cellpadding="0" cellspacing="0" style="background-color:#eef4fb; padding:30px 0;">
    <tr><td align="center">
      <table width="600" cellpadding="0" cellspacing="0"
             style="background-color:#ffffff; border-radius:8px; overflow:hidden; box-shadow:0 2px 8px rgba(13,59,102,0.15);">
        <tr>
          <td style="background-color:#0d3b66; padding:24px 32px;">
            <span style="color:#ffffff; font-size:20px; font-weight:600;">{company_name}</span><br>
            <span style="color:#a9c6e8; font-size:13px;">Gestão da Qualidade de Fornecedores</span>
          </td>
        </tr>
        <tr>
          <td style="padding:32px;">
            <h2 style="color:#e08e0b; margin-top:0; font-family: 'Segoe UI', Arial, sans-serif;">Confirmação de Certificado</h2>
            <p style="color:#333333; font-size:15px; line-height:1.5;">
              Prezados, <strong>{supplier_name}</strong>,
            </p>
            <p style="color:#333333; font-size:15px; line-height:1.5;">
              Não localizamos em nosso cadastro um certificado válido de <strong>{cert_type_label}</strong>
              para este fornecedor.
            </p>
            <div style="background-color:#f0f6fc; border-left:4px solid #e08e0b; padding:14px 18px; margin:20px 0; border-radius:4px;">
              <p style="margin:0; color:#0d3b66; font-size:14px; line-height:1.6;">
                <strong>Se a empresa possui o certificado:</strong> responda este e-mail anexando o arquivo
                (cópia atualizada).<br><br>
                <strong>Se a empresa não possui o certificado:</strong> por favor, informe-nos na resposta,
                para que possamos seguir com o <em>Risk Assessment</em> do fornecedor.
              </p>
            </div>
            <p style="color:#333333; font-size:15px; line-height:1.5;">
              Em caso de dúvidas, entre em contato com nossa equipe de Qualidade de Fornecedores.
            </p>
            <p style="color:#333333; font-size:15px; line-height:1.5;">
              Atenciosamente,<br>Equipe de Qualidade de Fornecedores
            </p>
          </td>
        </tr>
        <tr>
          <td style="background-color:#0d3b66; padding:14px 32px;">
            <span style="color:#a9c6e8; font-size:11px;">
              Esta é uma mensagem automática do sistema de gestão de certificados de fornecedores.
            </span>
          </td>
        </tr>
      </table>
    </td></tr>
  </table>
</body>
</html>
"""


def _body_pt(supplier_name, cert_type_label, expiry_date_str, days_left, tier, company_name):
    if tier == NO_DATE:
        return _body_missing_pt(supplier_name, cert_type_label, company_name)
    if tier == EXPIRED:
        title = "Certificado Vencido"
        accent_color = "#c0392b"
        status_message = (
            f"O certificado está <strong>vencido há {abs(days_left)} dia(s)</strong>. "
            f"Solicitamos a regularização o quanto antes."
        )
    else:
        title = "Renovação de Certificado"
        accent_color = "#1f6fb2"
        status_message = f"O certificado vence em <strong>{days_left} dia(s)</strong>."

    return f"""\
<html>
<body style="font-family: 'Segoe UI', Arial, sans-serif; background-color:#eef4fb; margin:0; padding:0;">
  <table width="100%" cellpadding="0" cellspacing="0" style="background-color:#eef4fb; padding:30px 0;">
    <tr><td align="center">
      <table width="600" cellpadding="0" cellspacing="0"
             style="background-color:#ffffff; border-radius:8px; overflow:hidden; box-shadow:0 2px 8px rgba(13,59,102,0.15);">
        <tr>
          <td style="background-color:#0d3b66; padding:24px 32px;">
            <span style="color:#ffffff; font-size:20px; font-weight:600;">{company_name}</span><br>
            <span style="color:#a9c6e8; font-size:13px;">Gestão da Qualidade de Fornecedores</span>
          </td>
        </tr>
        <tr>
          <td style="padding:32px;">
            <h2 style="color:{accent_color}; margin-top:0; font-family: 'Segoe UI', Arial, sans-serif;">{title}</h2>
            <p style="color:#333333; font-size:15px; line-height:1.5;">
              Prezados, <strong>{supplier_name}</strong>,
            </p>
            <p style="color:#333333; font-size:15px; line-height:1.5;">
              Identificamos que o certificado de <strong>{cert_type_label}</strong> deste fornecedor
              tem validade até <strong>{expiry_date_str}</strong>.
            </p>
            <p style="color:#333333; font-size:15px; line-height:1.5;">{status_message}</p>
            <div style="background-color:#f0f6fc; border-left:4px solid {accent_color}; padding:14px 18px; margin:20px 0; border-radius:4px;">
              <p style="margin:0; color:#0d3b66; font-size:14px; line-height:1.5;">
                Solicitamos o envio da cópia atualizada do certificado renovado assim que possível,
                para manutenção do cadastro de fornecedores homologados.
              </p>
            </div>
            <p style="color:#333333; font-size:15px; line-height:1.5;">
              Em caso de dúvidas, entre em contato com nossa equipe de Qualidade de Fornecedores.
            </p>
            <p style="color:#333333; font-size:15px; line-height:1.5;">
              Atenciosamente,<br>Equipe de Qualidade de Fornecedores
            </p>
          </td>
        </tr>
        <tr>
          <td style="background-color:#0d3b66; padding:14px 32px;">
            <span style="color:#a9c6e8; font-size:11px;">
              Esta é uma mensagem automática do sistema de gestão de certificados de fornecedores.
            </span>
          </td>
        </tr>
      </table>
    </td></tr>
  </table>
</body>
</html>
"""


def build_email_body(supplier_name, cert_type_label, expiry_date_str, days_left, tier,
                      company_name, language="en"):
    """Monta o corpo do e-mail no idioma solicitado ('en' ou 'pt-br')."""
    if tier == NO_DATE:  # certificado em branco: sempre em português
        return _body_missing_pt(supplier_name, cert_type_label, company_name)
    if language == "pt-br":
        return _body_pt(supplier_name, cert_type_label, expiry_date_str, days_left, tier, company_name)
    return _body_en(supplier_name, cert_type_label, expiry_date_str, days_left, tier, company_name)
