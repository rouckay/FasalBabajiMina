<?php

namespace App\Http\Controllers;

use App\Models\CustomerNumeraha;
use App\Models\Numeraha;
use LaravelDaily\Invoices\Invoice;
use LaravelDaily\Invoices\Classes\Buyer;
use LaravelDaily\Invoices\Classes\InvoiceItem;
use App\Models\FilamentModel; // Use your actual model
use Illuminate\Http\Request;
use Barryvdh\DomPDF\Facade\Pdf;
use Elibyy\TCPDF\Facades\TCPDF;
use FontLib\Table\Type\name;
// use PDF;
use PhpOffice\PhpWord\PhpWord;
use PhpOffice\PhpWord\IOFactory;
use Carbon\Carbon;
use Dompdf\Dompdf;
use Dompdf\Options;
use PhpOffice\PhpWord\TemplateProcessor;

class SharwaliTarifa extends Controller
{
    public function downloadInvoice($id)
    {
        $filament = CustomerNumeraha::with(['customer', 'numeraha'])->find($id);

        // Handle potential null values in the item details
        $numeraha_id = $filament->numeraha_id ?? 'Unnamed Item';
        $customer_id = $filament->customer_id ?? 0; // Fallback price
        $customer = $filament->customer->name ?? 'Unknown';
        $father_name = $filament->customer->father_name ?? 'Unknown';
        $grand_father_name = $filament->customer->grand_father_name ?? 'Unknown';
        $province = $filament->customer->province ?? 'Unknown';
        $village = $filament->customer->village ?? 'Unknown';
        $tazkira = $filament->customer->tazkira ?? 'Unknown';
        $mobile_number = $filament->customer->mobile_number ?? 'Unknown';
        $job = $filament->customer->job ?? 'Unknown';
        $Customer_image = $filament->customer->Customer_image ?? 'Unknown';
        $responsable_name = $filament->customer->responsable_name ?? 'Unknown';
        $responsable_father_name = $filament->customer->responsable_father_name ?? 'Unknown';
        $responsable_grand_father_name = $filament->customer->responsable_grand_father_name ?? 'Unknown';
        $responsable_province = $filament->customer->responsable_province ?? 'Unknown';
        $responsable_village = $filament->customer->responsable_village ?? 'Unknown';
        $responsable_tazkira = $filament->customer->responsable_tazkira ?? 'Unknown';
        $responsable_mobile_number = $filament->customer->responsable_mobile_number ?? 'Unknown';
        $responsable_image = $filament->customer->responsable_image ?? 'Unknown';
        $responsable_job = $filament->customer->responsable_job ?? 'Unknown';
        $responsable_ = $filament->customer->responsable_ ?? 'Unknown';
        //numerah Detail is started here

        $numera_id = $filament->numeraha->numera_id ?? '0 AFG';
        $numera_type = $filament->numeraha->numera_type ?? '0 AFG';
        $Land_Area = $filament->numeraha->Land_Area ?? '0 AFG';
        $north = $filament->numeraha->north ?? '0';
        $south = $filament->numeraha->south ?? '0';
        $east = $filament->numeraha->east ?? '0';
        $west = $filament->numeraha->west ?? '0';
        $total_price = $filament->total_price ?? '0 AFG';
        $sharwali_tarifa_price = $filament->numeraha->sharwali_tarifa_price ?? '0 AFG';
        $date = Carbon::now();

        $pdf = new TCPDF();

        // Set document information
        $pdf::SetAuthor('Your Name');
        $pdf::SetTitle('Persian PDF');

        // Set RTL language support if needed
        $pdf::setRTL(true);

        // Add a page
        $pdf::AddPage();

        // Render Blade template into HTML with data
        $html = view('pdf', compact(
            'numeraha_id',
            'customer_id',
            'customer',
            'father_name',
            'grand_father_name',
            'province',
            'tazkira',
            'village',
            'tazkira',
            'mobile_number',
            'job',
            'Customer_image',
            'responsable_name',
            'responsable_father_name',
            'responsable_grand_father_name',
            'responsable_province',
            'responsable_village',
            'responsable_tazkira',
            'responsable_mobile_number',
            'responsable_image',
            'responsable_job',
            'numera_id',
            'numera_type',
            'Land_Area',
            'north',
            'south',
            'east',
            'west',
            'total_price',
            'sharwali_tarifa_price',
            'date'
        ))->render();

        // Pass the rendered HTML to TCPDF
        $pdf::writeHTML($html, true, false, true, false, '');

        // Output PDF as download
        return $pdf::Output('تعرفه.pdf', 'D');
    }
    public function downloadsoldDocs($id)
    {
        $filament = CustomerNumeraha::with(['customer', 'numeraha'])->find($id);

        // Handle potential null values in the item details
        $numeraha_id = $filament->numeraha_id ?? 'Unnamed Item';
        $customer_id = $filament->customer_id ?? 0; // Fallback price
        $customer = $filament->customer->name ?? 'Unknown';
        $father_name = $filament->customer->father_name ?? 'Unknown';
        $grand_father_name = $filament->customer->grand_father_name ?? 'Unknown';
        $province = $filament->customer->province ?? 'Unknown';
        $village = $filament->customer->village ?? 'Unknown';
        $tazkira = $filament->customer->tazkira ?? 'Unknown';
        $mobile_number = $filament->customer->mobile_number ?? 'Unknown';
        $job = $filament->customer->job ?? 'Unknown';
        $Customer_image = $filament->customer->Customer_image ?? 'Unknown';
        $responsable_name = $filament->customer->responsable_name ?? 'Unknown';
        $responsable_father_name = $filament->customer->responsable_father_name ?? 'Unknown';
        $responsable_grand_father_name = $filament->customer->responsable_grand_father_name ?? 'Unknown';
        $responsable_province = $filament->customer->responsable_province ?? 'Unknown';
        $responsable_village = $filament->customer->responsable_village ?? 'Unknown';
        $responsable_tazkira = $filament->customer->responsable_tazkira ?? 'Unknown';
        $responsable_mobile_number = $filament->customer->responsable_mobile_number ?? 'Unknown';
        $responsable_image = $filament->customer->responsable_image ?? 'Unknown';
        $responsable_job = $filament->customer->responsable_job ?? 'Unknown';
        $responsable_ = $filament->customer->responsable_ ?? 'Unknown';
        //numerah Detail is started here

        $numera_id = $filament->numeraha->numera_id ?? '0 AFG';
        $numera_type = $filament->numeraha->numera_type ?? '0 AFG';
        $Land_Area = $filament->numeraha->Land_Area ?? '0 AFG';
        $north = $filament->numeraha->north ?? '0';
        $south = $filament->numeraha->south ?? '0';
        $east = $filament->numeraha->east ?? '0';
        $west = $filament->numeraha->west ?? '0';
        $total_price = $filament->total_price ?? '0 AFG';
        $sharwali_tarifa_price = $filament->numeraha->sharwali_tarifa_price ?? '0 AFG';
        $date = Carbon::now();

        // Initialize TCPDF
        $pdf = new Dompdf();
        $pdf = new TCPDF();

        // Set document information
        $pdf::SetAuthor('Your Name');
        $pdf::SetTitle('Persian PDF');

        // Set RTL language support if needed
        $pdf::setRTL(true);

        // Add a page
        $pdf::AddPage();
        // $pdf->set_option('defaultFont', 'Courier');
        // Set custom fonts
        $pdf::SetFont('Courier', '', 12); // Use Times New Roman

        // Render Blade template into HTML with data
        $html = view('numerahDocs', compact(
            'numeraha_id',
            'customer_id',
            'customer',
            'father_name',
            'grand_father_name',
            'province',
            'tazkira',
            'village',
            'tazkira',
            'mobile_number',
            'job',
            'Customer_image',
            'responsable_name',
            'responsable_father_name',
            'responsable_grand_father_name',
            'responsable_province',
            'responsable_village',
            'responsable_tazkira',
            'responsable_mobile_number',
            'responsable_image',
            'responsable_job',
            'numera_id',
            'numera_type',
            'Land_Area',
            'north',
            'south',
            'east',
            'west',
            'total_price',
            'sharwali_tarifa_price',
            'date'
        ))->render();

        // Pass the rendered HTML to TCPDF
        $pdf::writeHTML($html, true, false, true, false, '');

        // Output PDF as download
        return $pdf::Output('د نمری سند.pdf', 'D');

    }
    public function downloadsoldWord($id)
    {
        // Load the data from the database
        $filament = CustomerNumeraha::with(['customer', 'numeraha'])->find($id);

        // Handle potential null values
        $numeraha_id = $filament->numeraha_id ?? 'Unnamed Item';
        $customer_id = $filament->customer_id ?? 0;
        $customer = $filament->customer->name ?? 'Unknown';
        $father_name = $filament->customer->father_name ?? 'Unknown';
        $grand_father_name = $filament->customer->grand_father_name ?? 'Unknown';
        $province = $filament->customer->province ?? 'Unknown';
        $village = $filament->customer->village ?? 'Unknown';
        $tazkira = $filament->customer->tazkira ?? 'Unknown';
        $mobile_number = $filament->customer->mobile_number ?? 'Unknown';
        $job = $filament->customer->job ?? 'Unknown';
        $Customer_image = $filament->customer->Customer_image ?? 'Unknown';
        $responsable_name = $filament->customer->responsable_name ?? 'Unknown';
        $responsable_father_name = $filament->customer->responsable_father_name ?? 'Unknown';
        $responsable_grand_father_name = $filament->customer->responsable_grand_father_name ?? 'Unknown';
        $responsable_province = $filament->customer->responsable_province ?? 'Unknown';
        $responsable_village = $filament->customer->responsable_village ?? 'Unknown';
        $responsable_tazkira = $filament->customer->responsable_tazkira ?? 'Unknown';
        $responsable_mobile_number = $filament->customer->responsable_mobile_number ?? 'Unknown';
        $responsable_image = $filament->customer->responsable_image ?? 'Unknown';
        $responsable_job = $filament->customer->responsable_job ?? 'Unknown';
        $numera_id = $filament->numeraha->numera_id ?? '0 AFG';
        $numera_type = $filament->numeraha->numera_type ?? 'Unknown';
        $Land_Area = $filament->numeraha->Land_Area ?? '0 AFG';
        $north = $filament->numeraha->north ?? '0';
        $south = $filament->numeraha->south ?? '0';
        $east = $filament->numeraha->east ?? '0';
        $west = $filament->numeraha->west ?? '0';
        $total_price = $filament->total_price ?? '0 AFG';
        $sharwali_tarifa_price = $filament->numeraha->sharwali_tarifa_price ?? '0 AFG';
        $date = Carbon::now()->toDateString();

        // Load the Word template
        $templateProcessor = new TemplateProcessor(storage_path('app/templates/numerah_template.docx'));

        // Replace placeholders with actual data
        $templateProcessor->setValue('numeraha_id', $numeraha_id);
        $templateProcessor->setValue('customer_id', $customer_id);
        $templateProcessor->setValue('customer', $customer);
        $templateProcessor->setValue('father_name', $father_name);
        $templateProcessor->setValue('grand_father_name', $grand_father_name);
        $templateProcessor->setValue('province', $province);
        $templateProcessor->setValue('village', $village);
        $templateProcessor->setValue('tazkira', $tazkira);
        $templateProcessor->setValue('mobile_number', $mobile_number);
        $templateProcessor->setValue('job', $job);
        $templateProcessor->setImageValue('Customer_image', [
            'path' => storage_path('app/public/' . $Customer_image),
            'width' => 100,
            'height' => 100,
            'ratio' => true,
        ]);

        $templateProcessor->setImageValue('responsable_image', [
            'path' => storage_path('app/public/' . $responsable_image),
            'width' => 100,
            'height' => 100,
            'ratio' => true,
        ]);

        $templateProcessor->setValue('responsable_name', $responsable_name);
        $templateProcessor->setValue('responsable_father_name', $responsable_father_name);
        $templateProcessor->setValue('responsable_grand_father_name', $responsable_grand_father_name);
        $templateProcessor->setValue('responsable_province', $responsable_province);
        $templateProcessor->setValue('responsable_village', $responsable_village);
        $templateProcessor->setValue('responsable_tazkira', $responsable_tazkira);
        $templateProcessor->setValue('responsable_mobile_number', $responsable_mobile_number);
        $templateProcessor->setValue('responsable_job', $responsable_job);
        $templateProcessor->setValue('numera_id', $numera_id);
        $templateProcessor->setValue('numera_type', $numera_type);
        $templateProcessor->setValue('Land_Area', $Land_Area);
        $templateProcessor->setValue('north', $north);
        $templateProcessor->setValue('south', $south);
        $templateProcessor->setValue('east', $east);
        $templateProcessor->setValue('west', $west);
        $templateProcessor->setValue('total_price', $total_price);
        $templateProcessor->setValue('sharwali_tarifa_price', $sharwali_tarifa_price);
        $templateProcessor->setValue('date', $date);

        // Save the resulting file and return as a download
        $fileName = 'numerah_sold_document.docx';
        $tempFile = storage_path("app/public/$fileName");
        $templateProcessor->saveAs($tempFile);

        // Return the file as a download response
        return response()->download($tempFile)->deleteFileAfterSend(true);
    }
}
