<?php
namespace App\Http\Controllers;

use App\Models\Parcel;
use Illuminate\Http\Request;
use App\Services\ParcelService;

class ParcelController extends Controller
{
    public function __construct(private ParcelService $service) {}
    // ← OOP: Constructor injection. The controller declares what it needs;
    //   Lumen's service container provides it automatically.

    public function index()
    {
        return $this->success($this->service->getAll());
    }

    public function show($id)
    {
        $parcel = $this->service->findOrFail($id);
        return $this->success($parcel);
    }

    public function store(Request $request)
    {
        $this->validate($request, [
            'tracking_no'    => 'required|string|max:64|unique:parcels,tracking_no',
            'recipient_name' => 'required|string|max:100',
            'address'        => 'required|string|max:255',
            'weight'         => 'nullable|numeric|min:0',
            'status'         => 'nullable|in:pending,in_transit,delivered',
        ]);
        $parcel = $this->service->create($request->all());
        return $this->success($parcel, 'OK', 201);
    }

    public function update(Request $request, $id)
    {
        log($id);
        $parcel = $this->service->findOrFail($id);
        $this->validate($request, [
            'tracking_no'    => 'sometimes|string|max:64|unique:parcels,tracking_no,' . $id,
            'recipient_name' => 'sometimes|string|max:100',
            'address'        => 'sometimes|string|max:255',
            'weight'         => 'nullable|numeric|min:0',
            'status'         => 'nullable|in:pending,in_transit,delivered',
        ]);
        $parcel = $this->service->update($parcel, $request->all());
        return $this->success($parcel);
    }

    public function destroy($id)
    {
        $parcel = $this->service->findOrFail($id);
        $this->service->softDelete($parcel);
        return $this->success(null, 'Deleted', 200);
    }

    public function deleted()
    {
        return $this->success($this->service->getDeleted());
    }
}
