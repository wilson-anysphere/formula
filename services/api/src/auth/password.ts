import bcrypt from "bcryptjs";

export async function hashPassword(password: string): Promise<string> {
  // Keep production hashing strong, but speed up vitest/pg-mem integration tests.
  const saltRounds = process.env.NODE_ENV === "test" ? 4 : 12;
  return bcrypt.hash(password, saltRounds);
}

export async function verifyPassword(password: string, passwordHash: string): Promise<boolean> {
  return bcrypt.compare(password, passwordHash);
}
