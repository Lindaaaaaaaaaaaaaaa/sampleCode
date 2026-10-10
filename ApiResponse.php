<?php
namespace App\Http\Helpers;
trait ApiResponse
    // ← OOP: Trait
{
    protected function success(        // ← OOP: protected visibility
        $data = null,
        string $message = 'OK',
        int $status = 200
    ) {
        return response()->json([
            'success' => true,
            'message' => $message,
            'data'    => $data,
        ], $status);
    }

    protected function error(string $message, int $status, array $errors = [])
    {
        $body = [
            'success' => false,
            'message' => $message,
            'data'    => null,
        ];
        if (!empty($errors)) {
            $body['errors'] = $errors;
        }
        return response()->json($body, $status);
    }
}
